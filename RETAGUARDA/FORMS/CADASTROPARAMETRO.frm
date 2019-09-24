VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROPARAMETRO 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetros Sistema"
   ClientHeight    =   6540
   ClientLeft      =   2280
   ClientTop       =   2460
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROPARAMETRO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10890
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   10
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Descritores"
      TabPicture(0)   =   "CADASTROPARAMETRO.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LISTADESCR"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdMATARDESCR"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLIMPARDESCR"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSAIRDESCR"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "List1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Forma Faturamento"
      TabPicture(1)   =   "CADASTROPARAMETRO.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label49(0)"
      Tab(1).Control(1)=   "Label48"
      Tab(1).Control(2)=   "Label47"
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(7)=   "lblNrCon(2)"
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(9)=   "Label5"
      Tab(1).Control(10)=   "Label4"
      Tab(1).Control(11)=   "Label50"
      Tab(1).Control(12)=   "Label58"
      Tab(1).Control(13)=   "Label59"
      Tab(1).Control(14)=   "LISTAVENDA"
      Tab(1).Control(15)=   "chkPermite_Desconto"
      Tab(1).Control(16)=   "txtCredito"
      Tab(1).Control(17)=   "txtDebito"
      Tab(1).Control(18)=   "chkContabiliza"
      Tab(1).Control(19)=   "txtPercJuros"
      Tab(1).Control(20)=   "txtDiasPrazo"
      Tab(1).Control(21)=   "cmbFORMA"
      Tab(1).Control(22)=   "txtParcela"
      Tab(1).Control(23)=   "cmdSAIRVENDA"
      Tab(1).Control(24)=   "cmdLIMPARVENDA"
      Tab(1).Control(25)=   "cmdMatarVENDA"
      Tab(1).Control(26)=   "txtDESCTIPOVENDA"
      Tab(1).Control(27)=   "txtTIPOVENDA"
      Tab(1).Control(28)=   "cmbCC"
      Tab(1).Control(29)=   "cmbAuxForma"
      Tab(1).Control(30)=   "cmbCCAux"
      Tab(1).Control(31)=   "chkPreFatura"
      Tab(1).Control(32)=   "chkPagar"
      Tab(1).Control(33)=   "chkReceber"
      Tab(1).Control(34)=   "chkPermiteParcelar"
      Tab(1).Control(35)=   "cmbADMCARTAO"
      Tab(1).Control(36)=   "cmbADMCartaoAUX"
      Tab(1).Control(37)=   "txtDiaVencto"
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "Forma &Pagamento"
      TabPicture(2)   =   "CADASTROPARAMETRO.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(2)=   "LISTAPAGTO"
      Tab(2).Control(3)=   "chkCBalcao"
      Tab(2).Control(4)=   "chkBaixaAuto"
      Tab(2).Control(5)=   "chkCTesoraria"
      Tab(2).Control(6)=   "chkPagto"
      Tab(2).Control(7)=   "cmdMATARPAGTO"
      Tab(2).Control(8)=   "cmdLIMPARPAGTO"
      Tab(2).Control(9)=   "cmdSAIRPAGTO"
      Tab(2).Control(10)=   "txtDESCFORMAPAGTO"
      Tab(2).Control(11)=   "txtFORMAPAGTO"
      Tab(2).Control(12)=   "chkFunc"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "&Entrada NFe"
      TabPicture(3)   =   "CADASTROPARAMETRO.frx":5C66
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblestoque"
      Tab(3).Control(1)=   "Label9"
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(3)=   "LISTAENTRADA"
      Tab(3).Control(4)=   "txtboleto"
      Tab(3).Control(5)=   "txtTIPOENTRADA"
      Tab(3).Control(6)=   "txtDESCTIPOENTRADA"
      Tab(3).Control(7)=   "cmdmatarentrada"
      Tab(3).Control(8)=   "cmdLIMPARentrada"
      Tab(3).Control(9)=   "cmdSAIRentrada"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Tributação"
      TabPicture(4)   =   "CADASTROPARAMETRO.frx":5C82
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Line1"
      Tab(4).Control(1)=   "lblicmsfora"
      Tab(4).Control(2)=   "lblsubst"
      Tab(4).Control(3)=   "lbliss"
      Tab(4).Control(4)=   "lblipi"
      Tab(4).Control(5)=   "lblicms"
      Tab(4).Control(6)=   "Label14"
      Tab(4).Control(7)=   "Label13"
      Tab(4).Control(8)=   "Label12"
      Tab(4).Control(9)=   "Label60(0)"
      Tab(4).Control(10)=   "Label60(1)"
      Tab(4).Control(11)=   "Label49(1)"
      Tab(4).Control(12)=   "Label49(2)"
      Tab(4).Control(13)=   "Label1(10)"
      Tab(4).Control(14)=   "Label1(9)"
      Tab(4).Control(15)=   "Label1(1)"
      Tab(4).Control(16)=   "Label1(2)"
      Tab(4).Control(17)=   "Label1(3)"
      Tab(4).Control(18)=   "Label49(3)"
      Tab(4).Control(19)=   "Label49(4)"
      Tab(4).Control(20)=   "Label61"
      Tab(4).Control(21)=   "Label1(4)"
      Tab(4).Control(22)=   "lstCFOP"
      Tab(4).Control(23)=   "txtICMS_Fora"
      Tab(4).Control(24)=   "txtSUBST"
      Tab(4).Control(25)=   "txtISS"
      Tab(4).Control(26)=   "txtIPI"
      Tab(4).Control(27)=   "txtICMS_Dentro"
      Tab(4).Control(28)=   "txtFISCO"
      Tab(4).Control(29)=   "cmdCFOPMATA"
      Tab(4).Control(30)=   "cmdCFOPLIMPA"
      Tab(4).Control(31)=   "cmdCFOPSAIR"
      Tab(4).Control(32)=   "txtDesc_CFOP"
      Tab(4).Control(33)=   "txtCFOP_ID"
      Tab(4).Control(34)=   "txtUFORIGEM"
      Tab(4).Control(35)=   "cmbUFDestino"
      Tab(4).Control(36)=   "txtCOFINS"
      Tab(4).Control(37)=   "txtPIS"
      Tab(4).Control(38)=   "cmbPIS"
      Tab(4).Control(39)=   "cmbCOFINS"
      Tab(4).Control(40)=   "cmbCSTICMS"
      Tab(4).Control(41)=   "txtBaseReduz"
      Tab(4).Control(42)=   "cmbCSTORIG"
      Tab(4).ControlCount=   43
      TabCaption(5)   =   "Família de Produtos"
      TabPicture(5)   =   "CADASTROPARAMETRO.frx":5C9E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "CFOPS padrões"
      TabPicture(6)   =   "CADASTROPARAMETRO.frx":5CBA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).Control(1)=   "Frame6"
      Tab(6).Control(2)=   "Frame5"
      Tab(6).Control(3)=   "Frame4"
      Tab(6).Control(4)=   "frame"
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "Balança"
      TabPicture(7)   =   "CADASTROPARAMETRO.frx":5CD6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label52"
      Tab(7).Control(1)=   "Line2(0)"
      Tab(7).Control(2)=   "Line2(1)"
      Tab(7).Control(3)=   "lblMSG(0)"
      Tab(7).Control(4)=   "lblMSG(1)"
      Tab(7).Control(5)=   "Line3"
      Tab(7).Control(6)=   "Line2(2)"
      Tab(7).Control(7)=   "Text2"
      Tab(7).Control(8)=   "cmdSairBalanca"
      Tab(7).Control(9)=   "cmdLimpaBalanca"
      Tab(7).Control(10)=   "chkPanific"
      Tab(7).Control(11)=   "txtTamanhoCodgProdBarra"
      Tab(7).Control(12)=   "Frame1"
      Tab(7).Control(13)=   "Text1"
      Tab(7).Control(14)=   "txtTamanhoPesoValorBarra"
      Tab(7).Control(15)=   "cmdGravarBalanca"
      Tab(7).Control(16)=   "txtCodgProdutoReserva"
      Tab(7).Control(17)=   "txtCasaInicioCodgProdBarra"
      Tab(7).Control(18)=   "Text4"
      Tab(7).ControlCount=   19
      TabCaption(8)   =   "CSOSN"
      TabPicture(8)   =   "CADASTROPARAMETRO.frx":5CF2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label53"
      Tab(8).Control(1)=   "Label54"
      Tab(8).Control(2)=   "Label55"
      Tab(8).Control(3)=   "adoCSOSN"
      Tab(8).Control(4)=   "txtCodgCSOSN"
      Tab(8).Control(5)=   "txtDescCSOSN"
      Tab(8).Control(6)=   "txtInstrucaoCSOSN"
      Tab(8).ControlCount=   7
      TabCaption(9)   =   "CST"
      TabPicture(9)   =   "CADASTROPARAMETRO.frx":5D0E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label56"
      Tab(9).Control(1)=   "Label57"
      Tab(9).Control(2)=   "adoCST"
      Tab(9).Control(3)=   "txtDescCST"
      Tab(9).Control(4)=   "txtCodgCST"
      Tab(9).ControlCount=   5
      Begin VB.ComboBox cmbCSTORIG 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -65160
         TabIndex        =   113
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtBaseReduz 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   360
         Left            =   -69840
         TabIndex        =   212
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbCSTICMS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73680
         TabIndex        =   108
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cmbCOFINS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -67800
         TabIndex        =   116
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cmbPIS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73680
         TabIndex        =   114
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtPIS 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   -70560
         MaxLength       =   6
         TabIndex        =   115
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtCOFINS 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   -65400
         MaxLength       =   6
         TabIndex        =   117
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox cmbUFDestino 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -65640
         TabIndex        =   107
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtUFORIGEM 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   360
         Left            =   -67440
         MaxLength       =   2
         TabIndex        =   106
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtDiaVencto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -69885
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox cmbADMCartaoAUX 
         BackColor       =   &H00FFC0C0&
         Height          =   360
         Left            =   -72120
         TabIndex        =   198
         Top             =   3180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbADMCARTAO 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   -72120
         TabIndex        =   196
         Top             =   3180
         Width           =   5415
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   855
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   195
         Text            =   "CADASTROPARAMETRO.frx":5D2A
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtCasaInicioCodgProdBarra 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   -72600
         MaxLength       =   8
         TabIndex        =   18
         Text            =   "0"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtCodgCST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -74865
         MaxLength       =   4
         TabIndex        =   190
         Top             =   1215
         Width           =   1095
      End
      Begin VB.TextBox txtDescCST 
         Appearance      =   0  'Flat
         Height          =   1320
         Left            =   -73665
         MultiLine       =   -1  'True
         TabIndex        =   189
         Top             =   1215
         Width           =   9135
      End
      Begin VB.TextBox txtInstrucaoCSOSN 
         Appearance      =   0  'Flat
         Height          =   2040
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   185
         Top             =   2760
         Width           =   9135
      End
      Begin VB.TextBox txtDescCSOSN 
         Appearance      =   0  'Flat
         Height          =   1320
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   184
         Top             =   1035
         Width           =   9135
      End
      Begin VB.TextBox txtCodgCSOSN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   183
         Top             =   1035
         Width           =   1095
      End
      Begin VB.TextBox txtCodgProdutoReserva 
         Height          =   375
         Left            =   -71640
         TabIndex        =   181
         Top             =   4740
         Width           =   1095
      End
      Begin VB.CheckBox chkPermiteParcelar 
         Caption         =   "Permite Parcelar?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -66600
         TabIndex        =   179
         ToolTipText     =   "Título é baixado no ato do fechamento da venda"
         Top             =   1980
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.CheckBox chkReceber 
         Caption         =   "Contas à Receber?"
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
         Left            =   -66600
         TabIndex        =   178
         Top             =   1260
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkPagar 
         Caption         =   "Contas à Pagar?"
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
         Left            =   -66600
         TabIndex        =   177
         Top             =   1500
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkPreFatura 
         Caption         =   "Pré Fatura?"
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
         Left            =   -66600
         TabIndex        =   176
         Top             =   1020
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox cmbCCAux 
         BackColor       =   &H00FFC0C0&
         Height          =   360
         Left            =   -72120
         TabIndex        =   175
         Top             =   2700
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbAuxForma 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   -69840
         TabIndex        =   174
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbCC 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   -72120
         TabIndex        =   12
         Top             =   2700
         Width           =   5415
      End
      Begin VB.CheckBox chkFunc 
         Caption         =   "Somente P/ Funcionário?"
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
         Left            =   -67920
         TabIndex        =   172
         ToolTipText     =   "Título é baixado no ato do fechamento da venda"
         Top             =   1500
         UseMaskColor    =   -1  'True
         Width           =   3135
      End
      Begin VB.CommandButton cmdGravarBalanca 
         Caption         =   "&Gravar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -72960
         Picture         =   "CADASTROPARAMETRO.frx":5D52
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   5460
         Width           =   900
      End
      Begin VB.TextBox txtTamanhoPesoValorBarra 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   -69240
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "5"
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   -71520
         MultiLine       =   -1  'True
         TabIndex        =   170
         Text            =   "CADASTROPARAMETRO.frx":7311
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Código de Barras Balança por"
         Height          =   735
         Left            =   -67440
         TabIndex        =   167
         Top             =   1800
         Width           =   3015
         Begin VB.OptionButton optValor 
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
            Height          =   240
            Left            =   2040
            TabIndex        =   169
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optGramas 
            Caption         =   "Gramas"
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
            TabIndex        =   168
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         ForeColor       =   &H00400000&
         Height          =   5550
         ItemData        =   "CADASTROPARAMETRO.frx":7354
         Left            =   7560
         List            =   "CADASTROPARAMETRO.frx":73B2
         MouseIcon       =   "CADASTROPARAMETRO.frx":76C0
         MousePointer    =   99  'Custom
         TabIndex        =   137
         Top             =   780
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   133
         Top             =   660
         Width           =   7335
         Begin VB.TextBox txtTipo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            MaxLength       =   2
            TabIndex        =   0
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   5415
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   840
            TabIndex        =   1
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   225
            Left            =   1800
            TabIndex        =   135
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Codigo:"
            Height          =   225
            Left            =   840
            TabIndex        =   134
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtTIPOVENDA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -73800
         MaxLength       =   8
         TabIndex        =   3
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtDESCTIPOVENDA 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -71760
         MaxLength       =   50
         TabIndex        =   4
         Top             =   780
         Width           =   5055
      End
      Begin VB.TextBox txtFORMAPAGTO 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   -71880
         MaxLength       =   8
         TabIndex        =   132
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtDESCFORMAPAGTO 
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
         Height          =   360
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   131
         Top             =   1860
         Width           =   5895
      End
      Begin VB.CommandButton cmdMatarVENDA 
         Caption         =   "&Excluir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65400
         Picture         =   "CADASTROPARAMETRO.frx":79CA
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   4500
         Width           =   900
      End
      Begin VB.CommandButton cmdLIMPARVENDA 
         Caption         =   "&Limpar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65400
         Picture         =   "CADASTROPARAMETRO.frx":8085
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   3540
         Width           =   900
      End
      Begin VB.CommandButton cmdSAIRVENDA 
         Caption         =   "&Voltar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65400
         Picture         =   "CADASTROPARAMETRO.frx":8687
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   5460
         Width           =   900
      End
      Begin VB.CommandButton cmdSAIRPAGTO 
         Caption         =   "&Voltar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65400
         Picture         =   "CADASTROPARAMETRO.frx":9811
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   5460
         Width           =   900
      End
      Begin VB.CommandButton cmdLIMPARPAGTO 
         Caption         =   "L&impar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65400
         Picture         =   "CADASTROPARAMETRO.frx":A99B
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   3540
         Width           =   900
      End
      Begin VB.CommandButton cmdMATARPAGTO 
         Caption         =   "&Excluir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65400
         Picture         =   "CADASTROPARAMETRO.frx":AF9D
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   4500
         Width           =   900
      End
      Begin VB.CommandButton cmdSAIRDESCR 
         Caption         =   "&Sair"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         Picture         =   "CADASTROPARAMETRO.frx":B658
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   5520
         Width           =   900
      End
      Begin VB.CommandButton cmdLIMPARDESCR 
         Caption         =   "&Limpar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   2040
         Picture         =   "CADASTROPARAMETRO.frx":BD28
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   5535
         Width           =   900
      End
      Begin VB.CommandButton cmdMATARDESCR 
         Caption         =   "&Excluir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1080
         Picture         =   "CADASTROPARAMETRO.frx":C32A
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   5535
         Width           =   900
      End
      Begin VB.CommandButton cmdSAIRentrada 
         Caption         =   "&Voltar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65760
         Picture         =   "CADASTROPARAMETRO.frx":C9E5
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   5100
         Width           =   900
      End
      Begin VB.CommandButton cmdLIMPARentrada 
         Caption         =   "Li&mpar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65760
         Picture         =   "CADASTROPARAMETRO.frx":DB6F
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   3060
         Width           =   900
      End
      Begin VB.CommandButton cmdmatarentrada 
         Caption         =   "&Excluir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65760
         Picture         =   "CADASTROPARAMETRO.frx":E171
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   4080
         Width           =   900
      End
      Begin VB.TextBox txtDESCTIPOENTRADA 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72120
         MaxLength       =   50
         TabIndex        =   109
         Top             =   1260
         Width           =   5775
      End
      Begin VB.TextBox txtTIPOENTRADA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72120
         MaxLength       =   8
         TabIndex        =   104
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtParcela 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -73800
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
      Begin VB.ComboBox cmbFORMA 
         Height          =   360
         Left            =   -71760
         TabIndex        =   5
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txtDiasPrazo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -72000
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCFOP_ID 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   103
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtDesc_CFOP 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -73680
         TabIndex        =   105
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton cmdCFOPSAIR 
         Caption         =   "&Voltar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65160
         Picture         =   "CADASTROPARAMETRO.frx":E82C
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   5520
         Width           =   900
      End
      Begin VB.CommandButton cmdCFOPLIMPA 
         Caption         =   "&Limpar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65160
         Picture         =   "CADASTROPARAMETRO.frx":F9B6
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   3600
         Width           =   900
      End
      Begin VB.CommandButton cmdCFOPMATA 
         Caption         =   "&Excluir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -65160
         Picture         =   "CADASTROPARAMETRO.frx":FFB8
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txtFISCO 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -73680
         TabIndex        =   121
         Top             =   3000
         Width           =   9135
      End
      Begin VB.TextBox txtPercJuros 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -68280
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   67
         Top             =   600
         Width           =   10455
         Begin VB.TextBox txtPercVenda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   8400
            MaxLength       =   4
            TabIndex        =   77
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtUN 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   74
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtDescGrupo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   73
            Top             =   240
            Width           =   6735
         End
         Begin VB.TextBox txtCodgGrupo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   72
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtDescUnidade 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   76
            Top             =   840
            Width           =   2775
         End
         Begin VB.CommandButton cmdGrupoSair 
            Caption         =   "&Voltar"
            CausesValidation=   0   'False
            Height          =   900
            Left            =   9480
            Picture         =   "CADASTROPARAMETRO.frx":10673
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   2130
            Width           =   900
         End
         Begin VB.CommandButton cmdGrupoLimpa 
            Caption         =   "&Limpar"
            CausesValidation=   0   'False
            Height          =   900
            Left            =   9480
            Picture         =   "CADASTROPARAMETRO.frx":117FD
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   240
            Width           =   900
         End
         Begin VB.CommandButton cmdGrupoMata 
            Caption         =   "&Excluir"
            CausesValidation=   0   'False
            Height          =   900
            Left            =   9480
            Picture         =   "CADASTROPARAMETRO.frx":11DFF
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   1185
            Width           =   900
         End
         Begin VB.CheckBox chkProducao 
            Caption         =   "É familia de produção?"
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
            Left            =   6720
            TabIndex        =   68
            Top             =   1250
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid grdFamilia 
            Bindings        =   "CADASTROPARAMETRO.frx":124BA
            Height          =   4095
            Left            =   120
            TabIndex        =   75
            Top             =   1560
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   18
            WrapCellPointer =   -1  'True
            RowDividerStyle =   3
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc adoFamilia 
            Height          =   330
            Left            =   4800
            Top             =   1440
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
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
            Caption         =   "Grid Cabeça"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%CompoePreçoVenda ="
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
            Left            =   5430
            TabIndex        =   180
            Top             =   840
            Width           =   2820
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Left            =   480
            TabIndex        =   79
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Unidade:"
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
            Left            =   360
            TabIndex        =   78
            Top             =   840
            Width           =   1050
         End
      End
      Begin VB.TextBox txtICMS_Dentro 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   360
         Left            =   -70560
         TabIndex        =   110
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtIPI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   66
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtISS 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   -65040
         TabIndex        =   65
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSUBST 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   360
         Left            =   -73320
         TabIndex        =   119
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtICMS_Fora 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   -67440
         TabIndex        =   112
         Top             =   1440
         Width           =   495
      End
      Begin VB.Frame frame 
         Caption         =   "Redução de Impostos"
         ForeColor       =   &H00400000&
         Height          =   1335
         Left            =   -74880
         TabIndex        =   48
         Top             =   5100
         Width           =   10455
         Begin VB.TextBox txtdecont 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2160
            TabIndex        =   92
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtdemaq 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2160
            TabIndex        =   93
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtfemaq 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2160
            TabIndex        =   94
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtfeapa 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2160
            TabIndex        =   95
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtdencont 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9240
            TabIndex        =   96
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtdenmaq 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9240
            TabIndex        =   97
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtfenmaq 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9240
            TabIndex        =   98
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtfenapa 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9240
            TabIndex        =   99
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblde_c 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DE Contribuinte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   64
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DE N. Contribuinte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7890
            TabIndex        =   63
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DE Maq. Imp. Cont."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   62
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DE Maq. Imp. N. Cont."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7650
            TabIndex        =   61
            Top             =   480
            Width           =   1530
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FE Maq. Imp. Cont."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   60
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FE Maq. Imp. N. Cont."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7650
            TabIndex        =   59
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FE Apa Ind. Cont."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   58
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FE Apa Ind. N Cont."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7800
            TabIndex        =   57
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   3000
            TabIndex        =   56
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   3000
            TabIndex        =   55
            Top             =   480
            Width           =   180
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   3000
            TabIndex        =   54
            Top             =   720
            Width           =   180
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   3000
            TabIndex        =   53
            Top             =   960
            Width           =   180
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   10080
            TabIndex        =   52
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   10080
            TabIndex        =   51
            Top             =   480
            Width           =   180
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   10080
            TabIndex        =   50
            Top             =   720
            Width           =   180
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Left            =   10080
            TabIndex        =   49
            Top             =   960
            Width           =   180
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Saídas"
         ForeColor       =   &H00400000&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   45
         Top             =   1800
         Width           =   10455
         Begin VB.ComboBox cmbcfopsd 
            Height          =   360
            Left            =   1680
            TabIndex        =   82
            Top             =   240
            Width           =   8685
         End
         Begin VB.ComboBox cmbcfopsf 
            Height          =   360
            Left            =   1680
            TabIndex        =   83
            Top             =   600
            Width           =   8685
         End
         Begin VB.Label lblsaida 
            Caption         =   "Dentro do Estado"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblsaidaf 
            Caption         =   "Fora do Estado"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Entradas"
         ForeColor       =   &H00400000&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   42
         Top             =   720
         Width           =   10455
         Begin VB.ComboBox cmbcfoped 
            Height          =   360
            Left            =   1680
            TabIndex        =   80
            Top             =   240
            Width           =   8685
         End
         Begin VB.ComboBox cmbcfopef 
            Height          =   360
            Left            =   1680
            TabIndex        =   81
            Top             =   600
            Width           =   8685
         End
         Begin VB.Label Label37 
            Caption         =   "Fora do Estado"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label38 
            Caption         =   "Dentro do Estado"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Devoluções"
         ForeColor       =   &H00400000&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   35
         Top             =   2910
         Width           =   10455
         Begin VB.ComboBox cmbdvsd 
            Height          =   360
            Left            =   1680
            TabIndex        =   84
            Top             =   240
            Width           =   3735
         End
         Begin VB.ComboBox cmbdvsf 
            Height          =   360
            Left            =   6360
            TabIndex        =   86
            Top             =   240
            Width           =   4005
         End
         Begin VB.ComboBox cmbdved 
            Height          =   360
            Left            =   1680
            TabIndex        =   85
            Top             =   600
            Width           =   3735
         End
         Begin VB.ComboBox cmbdvef 
            Height          =   360
            Left            =   6360
            TabIndex        =   87
            Top             =   600
            Width           =   4005
         End
         Begin VB.Label Label39 
            Caption         =   "D.Est."
            Height          =   255
            Left            =   960
            TabIndex        =   41
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label40 
            Caption         =   "D.Est."
            Height          =   255
            Left            =   960
            TabIndex        =   40
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label42 
            Caption         =   "F.Est."
            Height          =   255
            Left            =   5820
            TabIndex        =   39
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label43 
            Caption         =   "F.Est."
            Height          =   255
            Left            =   5820
            TabIndex        =   38
            Top             =   270
            Width           =   615
         End
         Begin VB.Label lblsaidadev 
            Caption         =   "Saída"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblentdev 
            Caption         =   "Entrada"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Transferências"
         ForeColor       =   &H00400000&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   28
         Top             =   3990
         Width           =   10455
         Begin VB.ComboBox cmbtrsd 
            Height          =   360
            Left            =   1680
            TabIndex        =   88
            Top             =   240
            Width           =   3735
         End
         Begin VB.ComboBox cmbtrsf 
            Height          =   360
            Left            =   6360
            TabIndex        =   90
            Top             =   240
            Width           =   4005
         End
         Begin VB.ComboBox cmbtred 
            Height          =   360
            Left            =   1680
            TabIndex        =   89
            Top             =   600
            Width           =   3735
         End
         Begin VB.ComboBox cmbtref 
            Height          =   360
            Left            =   6360
            TabIndex        =   91
            Top             =   600
            Width           =   4005
         End
         Begin VB.Label Label23 
            Caption         =   "Entrada"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label33 
            Caption         =   "Saída"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label41 
            Caption         =   "F.Est."
            Height          =   255
            Left            =   5790
            TabIndex        =   32
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label44 
            Caption         =   "F.Est."
            Height          =   255
            Left            =   5790
            TabIndex        =   31
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label45 
            Caption         =   "D.Est."
            Height          =   255
            Left            =   960
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label46 
            Caption         =   "D.Est."
            Height          =   255
            Left            =   960
            TabIndex        =   29
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.TextBox txtboleto 
         Height          =   375
         Left            =   -72120
         TabIndex        =   27
         Top             =   1740
         Width           =   495
      End
      Begin VB.CheckBox chkContabiliza 
         Caption         =   "Gerar Fatura"
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
         Left            =   -66600
         TabIndex        =   26
         Top             =   1740
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtDebito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -72120
         MaxLength       =   6
         TabIndex        =   10
         Top             =   2220
         Width           =   975
      End
      Begin VB.TextBox txtCredito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   -68280
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2220
         Width           =   975
      End
      Begin VB.CheckBox chkPagto 
         Caption         =   "Ativa"
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
         Left            =   -70920
         TabIndex        =   25
         Top             =   1140
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCTesoraria 
         Caption         =   "Contabiliza Caixa Tesouraria?"
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
         Left            =   -67920
         TabIndex        =   24
         ToolTipText     =   "Soma no caixa Tesouraria"
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   3495
      End
      Begin VB.CheckBox chkBaixaAuto 
         Caption         =   "Baixa Automática?"
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
         Left            =   -67920
         TabIndex        =   23
         ToolTipText     =   "Título é baixado no ato do fechamento da venda"
         Top             =   1260
         UseMaskColor    =   -1  'True
         Width           =   3495
      End
      Begin VB.CheckBox chkCBalcao 
         Caption         =   "Contabiliza Caixa Balcão?"
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
         Left            =   -67920
         TabIndex        =   22
         ToolTipText     =   "Soma no caixa Tesouraria"
         Top             =   1020
         UseMaskColor    =   -1  'True
         Width           =   3495
      End
      Begin VB.CheckBox chkPermite_Desconto 
         Caption         =   "Permite Desconto?"
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
         Left            =   -66600
         TabIndex        =   21
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox txtTamanhoCodgProdBarra 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   -72600
         MaxLength       =   8
         TabIndex        =   19
         Text            =   "5"
         Top             =   2820
         Width           =   735
      End
      Begin VB.CheckBox chkPanific 
         Caption         =   "Panificadora?"
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
         Left            =   -66240
         TabIndex        =   17
         ToolTipText     =   "Soma no caixa Tesouraria"
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdLimpaBalanca 
         Caption         =   "L&impar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -73920
         Picture         =   "CADASTROPARAMETRO.frx":124D3
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5460
         Width           =   900
      End
      Begin VB.CommandButton cmdSairBalanca 
         Cancel          =   -1  'True
         Caption         =   "&Voltar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -74880
         Picture         =   "CADASTROPARAMETRO.frx":12AD5
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5460
         Width           =   900
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   1095
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "CADASTROPARAMETRO.frx":13C5F
         Top             =   2460
         Width           =   2175
      End
      Begin MSComctlLib.ListView LISTAVENDA 
         Height          =   2745
         Left            =   -74880
         TabIndex        =   138
         Top             =   3660
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   4842
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1411
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
            Object.Width           =   1411
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Juros"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "GeraFaturamento"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PermiteDesconto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "CentroCusto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "PreFatura"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "À Pagar"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "À Receber"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Parcelar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ADM Cartão"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView LISTAPAGTO 
         Height          =   4065
         Left            =   -74880
         TabIndex        =   139
         Top             =   2340
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   7170
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Situação"
            Object.Width           =   1762
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ContabBalcao"
            Object.Width           =   1960
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "BaixaAuto"
            Object.Width           =   1960
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ContabTesouraria"
            Object.Width           =   1960
         EndProperty
      End
      Begin MSComctlLib.ListView LISTAENTRADA 
         Height          =   3705
         Left            =   -74400
         TabIndex        =   140
         Top             =   2280
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   6535
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
      Begin MSComctlLib.ListView lstCFOP 
         Height          =   2925
         Left            =   -74955
         TabIndex        =   141
         Top             =   3480
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   5159
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CFOP"
            Object.Width           =   1960
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Origem"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CST.ICMS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ICMS Dentro"
            Object.Width           =   3919
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Destino"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "ICMS Fora"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "CST.PIS"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Pis"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "CST.COFINS"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Cofins"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "%Reduc.Base"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView LISTADESCR 
         Height          =   3705
         Left            =   120
         TabIndex        =   142
         Top             =   1680
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   6535
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   195987
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoCSOSN 
         Height          =   330
         Left            =   -73680
         Top             =   4980
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "CSON"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc adoCST 
         Height          =   330
         Left            =   -73665
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "CST"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Orig.CST:"
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
         Index           =   4
         Left            =   -66180
         TabIndex        =   213
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label61 
         Caption         =   "% Redução Base Cálculo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   -72360
         TabIndex        =   211
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label49 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   -66840
         TabIndex        =   210
         Top             =   1440
         Width           =   150
      End
      Begin VB.Label Label49 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   -69960
         TabIndex        =   209
         Top             =   1440
         Width           =   150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CST ICMS:"
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
         Index           =   3
         Left            =   -74790
         TabIndex        =   208
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CST COFINS:"
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
         Height          =   255
         Index           =   2
         Left            =   -69240
         TabIndex        =   207
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CST PIS:"
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
         Index           =   1
         Left            =   -74595
         TabIndex        =   206
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alq.PIS:"
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
         Height          =   255
         Index           =   9
         Left            =   -71520
         TabIndex        =   205
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alq.COFINS:"
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
         Index           =   10
         Left            =   -66675
         TabIndex        =   204
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label Label49 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   -69720
         TabIndex        =   203
         Top             =   1920
         Width           =   150
      End
      Begin VB.Label Label49 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   -64560
         TabIndex        =   202
         Top             =   1920
         Width           =   150
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "UF Destino:"
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
         Index           =   1
         Left            =   -66840
         TabIndex        =   201
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "UF Origem:"
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
         Left            =   -68640
         TabIndex        =   200
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DiasVencto:"
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
         Left            =   -71025
         TabIndex        =   199
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label58 
         Caption         =   "Administradora Cartão D/C:"
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
         Left            =   -74835
         TabIndex        =   197
         Top             =   3120
         Width           =   2595
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   -75000
         X2              =   -64320
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line3 
         X1              =   -71760
         X2              =   -71760
         Y1              =   1320
         Y2              =   3720
      End
      Begin VB.Label lblMSG 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Configurações Gerais"
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
         Index           =   1
         Left            =   -75000
         TabIndex        =   194
         Top             =   4080
         Width           =   10770
      End
      Begin VB.Label lblMSG 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Configurações por Estabelecimento"
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
         Index           =   0
         Left            =   -75000
         TabIndex        =   193
         Top             =   840
         Width           =   10770
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   -75000
         X2              =   -64320
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   -75000
         X2              =   -64320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   -74880
         TabIndex        =   192
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   -73620
         TabIndex        =   191
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Instrução:"
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
         Left            =   -73695
         TabIndex        =   188
         Top             =   2520
         Width           =   930
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   -73635
         TabIndex        =   187
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   -74895
         TabIndex        =   186
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Reserva Sequencia Código Produto:"
         Height          =   240
         Left            =   -74880
         TabIndex        =   182
         Top             =   4740
         Width           =   3120
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Centro de Custo:"
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
         Left            =   -73800
         TabIndex        =   173
         Top             =   2700
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Código:"
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
         TabIndex        =   166
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrição: "
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
         Left            =   -72840
         TabIndex        =   165
         Top             =   780
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Código Forma Pagamento:"
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
         Left            =   -74550
         TabIndex        =   164
         Top             =   1140
         Width           =   2565
      End
      Begin VB.Label Label7 
         Caption         =   "Descrição Forma Pagamento:"
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
         Left            =   -74760
         TabIndex        =   163
         Top             =   1620
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição Tipo Entrada:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   162
         Top             =   1260
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo Tipo Entrada:"
         Height          =   255
         Left            =   -73920
         TabIndex        =   161
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Parcela(s):"
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
         Left            =   -74880
         TabIndex        =   160
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblNrCon 
         Alignment       =   1  'Right Justify
         Caption         =   "PAGTO:"
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
         Left            =   -72600
         TabIndex        =   159
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DiasPrazo:"
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
         Left            =   -73035
         TabIndex        =   158
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Left            =   -74760
         TabIndex        =   157
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   -73680
         TabIndex        =   156
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Msg.Fisco:"
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
         Left            =   -74760
         TabIndex        =   155
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   -67200
         TabIndex        =   154
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Juros:"
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
         Left            =   -69000
         TabIndex        =   153
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblicms 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alq.ICMS Dentro UF:"
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
         Left            =   -72600
         TabIndex        =   152
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblipi 
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
         Height          =   255
         Left            =   -66840
         TabIndex        =   151
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbliss 
         AutoSize        =   -1  'True
         Caption         =   "%ISS:"
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
         Left            =   -65640
         TabIndex        =   150
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblsubst 
         Caption         =   "% ICMS Subst.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   149
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblicmsfora 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alq.ICMS Fora UF:"
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
         Left            =   -69285
         TabIndex        =   148
         Top             =   1440
         Width           =   1740
      End
      Begin VB.Label lblestoque 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emite Boleto S/N?:"
         Height          =   255
         Left            =   -73920
         TabIndex        =   147
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Venda Cartão Débito:"
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
         Left            =   -74235
         TabIndex        =   146
         Top             =   2220
         Width           =   2040
      End
      Begin VB.Label Label47 
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
         Height          =   240
         Left            =   -71040
         TabIndex        =   145
         Top             =   2220
         Width           =   150
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Venda Cartão Crédito:"
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
         Left            =   -70515
         TabIndex        =   144
         Top             =   2220
         Width           =   2115
      End
      Begin VB.Label Label49 
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
         Height          =   255
         Index           =   0
         Left            =   -67200
         TabIndex        =   143
         Top             =   2280
         Width           =   255
      End
      Begin VB.Line Line1 
         X1              =   -75000
         X2              =   -64200
         Y1              =   3420
         Y2              =   3420
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
      DesignWidth     =   10890
      DesignHeight    =   6540
   End
End
Attribute VB_Name = "frmCADASTROPARAMETRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim TabCFOP       As New ADODB.Recordset
   Dim FAMILIA_ID_N  As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   Me.Caption = Me.Caption & " - " & Me.Name

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select * from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status = 1"
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      If USUARIO_ID_N <> 144 Then
         If Not TabUSU!TIPO >= 4 Then
            'SSTab1.TabVisible(3) = False
            'SSTab1.TabVisible(4) = False
            'SSTab1.TabVisible(6) = False
            Else
              SSTab1.TabVisible(3) = True
              SSTab1.TabVisible(4) = True
              SSTab1.TabVisible(6) = True
         End If
         Else
            'SSTab1.TabVisible(3) = False
            'SSTab1.TabVisible(4) = False
            'SSTab1.TabVisible(6) = False
         End If
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close
   
   SETA_GRID_PAGTO

   cmbCCAux.Clear
   cmbCC.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = 'O'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCC.AddItem Trim(TabTemp!DESCRICAO)
      cmbCCAux.AddItem TabTemp!Codigo
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbADMCartaoAUX.Clear
   cmbADMCARTAO.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from CARTAOADM WITH (NOLOCK)"
   SQL = SQL & " where status = 'A'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbADMCARTAO.AddItem Trim(TabTemp.Fields("fantasia").Value)
      cmbADMCartaoAUX.AddItem TabTemp.Fields("cartaoadm_id").Value
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SSTab1.TabVisible(6) = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub lstCFOP_DblClick()
On Error Resume Next

   If Trim(lstCFOP.SelectedItem.Text) <> "" Then _
      MOSTRA_CFOP_CFOPUF_ALIQUOTA_UF_DO_GRID lstCFOP.SelectedItem.Text

Err.Clear
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'On Error GoTo ERRO_TRATA

   'If SSTab1.Tab = 0 Then _
      txtTIPO.SetFocus

   If SSTab1.Tab = 1 Then     'forma faturamento
      txtTIPOVENDA.SetFocus
      SETA_GRID_VENDA
   End If
   If SSTab1.Tab = 2 Then
      txtFORMAPAGTO.SetFocus
      SETA_GRID_PAGTO
   End If
   If SSTab1.Tab = 3 Then
      txtTIPOENTRADA.SetFocus
      SETA_GRID_ENTRADA
   End If
   If SSTab1.Tab = 4 Then
      If Trim(UF_EMPRESA_A) = "" Then _
         PEGA_DADOS_EMPRESA
   
      txtUFORIGEM.Text = "" & UF_EMPRESA_A
   
      cmbUFDestino.Clear
   
      SQL = "SELECT distinct(ESTADO) + ' - ' + descricao FROM UF WITH (NOLOCK)"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         cmbUFDestino.AddItem Trim(TabTemp.Fields(0).Value)
         TabTemp.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close
      
      cmbUFDestino.Text = "" & UF_EMPRESA_A
   
      cmbCSTICMS.Clear
   
      SQL = "SELECT codigo + ' - ' + descricao FROM CST WITH (NOLOCK)"
      SQL = SQL & " where tipo = 'ICMS'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         cmbCSTICMS.AddItem Trim(TabTemp.Fields(0).Value)
         TabTemp.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close

      cmbCSTORIG.Clear
      cmbCSTORIG.AddItem "00"
      cmbCSTORIG.AddItem "300"

'==========
      'ABRE_BANCO_GLOBAL
      'PIS
      cmbPIS.Clear

      'SQL = "SELECT right(MFTCODFIS,3) + ' - ' + mftdescri FROM MFTCLASFISTRI WITH (NOLOCK)"
      'SQL = SQL & " where left(mftdescri,3) = 'PIS'"
      'TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
      'While Not TabTemp.EOF
      '   cmbPIS.AddItem Trim(TabTemp.Fields(0).Value)
      '   TabTemp.MoveNext
      'Wend
      'If TabTemp.State = 1 Then _
         TabTemp.Close

      'COFINS
      cmbCOFINS.Clear

      'SQL = "SELECT right(MFTCODFIS,3) + ' - ' + mftdescri FROM MFTCLASFISTRI WITH (NOLOCK)"
      'SQL = SQL & " where left(mftdescri,6) = 'COFINS'"
      'TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
      'While Not TabTemp.EOF
      '   cmbCOFINS.AddItem Trim(TabTemp.Fields(0).Value)
      '   TabTemp.MoveNext
      'Wend
      'If TabTemp.State = 1 Then _
         TabTemp.Close
      'If CONECTA_GLOBAL.State = 1 Then _
         CONECTA_GLOBAL.Close

      SQL = "SELECT codigo + ' - ' + descricao FROM CST WITH (NOLOCK)"
      SQL = SQL & " where tipo = 'PISCOFINS'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabTemp.EOF
         cmbPIS.AddItem Trim(TabTemp.Fields(0).Value)
         cmbCOFINS.AddItem Trim(TabTemp.Fields(0).Value)
         TabTemp.MoveNext
      Wend
      If TabTemp.State = 1 Then _
         TabTemp.Close


      SETA_GRID_CFOP
      txtCFOP_ID.SetFocus
   End If
   If SSTab1.Tab = 5 Then
      SETA_GRID_GRUPO_PRODUTOS
      txtCodgGrupo.Enabled = True
      Frame3.Enabled = True
 '     txtCodgGrupo.SetFocus
   End If
   If SSTab1.Tab = 6 Then
      preencheComboCfop cmbcfopsd
      preencheComboCfop cmbcfopsf
      preencheComboCfop cmbcfoped
      preencheComboCfop cmbcfopef
      preencheComboCfop cmbdvsd
      preencheComboCfop cmbdvsf
      preencheComboCfop cmbdved
      preencheComboCfop cmbdvef
      preencheComboCfop cmbtrsd
      preencheComboCfop cmbtrsf
      preencheComboCfop cmbtred
      preencheComboCfop cmbtref

      MOSTRA_PERC_CONT

      cmbcfopsd.SetFocus
   End If
   If SSTab1.Tab = 7 Then _
      MOSTRA_BALANCA
   If SSTab1.Tab = 8 Then _
      SETA_GRID_CSOSN
   If SSTab1.Tab = 9 Then _
      SETA_GRID_CST

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSTab1_Click"
End Sub

Private Sub cmbCOFINS_GotFocus()
   cmbCOFINS.SelStart = 0
   cmbCOFINS.SelLength = Len(cmbCOFINS.Text)
   cmbCOFINS.BackColor = &HC0FFFF
End Sub

Private Sub cmbCSTICMS_GotFocus()
   cmbCSTICMS.SelStart = 0
   cmbCSTICMS.SelLength = Len(cmbCSTICMS.Text)
   cmbCSTICMS.BackColor = &HC0FFFF
End Sub

Private Sub cmbCSTICMS_Click()
   txtICMS_Dentro.SetFocus
End Sub

Private Sub cmbcsticms_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtICMS_Dentro.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcsticms_KeyPress"
End Sub

Private Sub cmbCSTICMS_LostFocus()
   cmbCSTICMS.BackColor = &HFFFFFF
End Sub

Private Sub cmbCSTORIG_GotFocus()
   cmbCSTORIG.SelStart = 0
   cmbCSTORIG.SelLength = Len(cmbCSTORIG.Text)
   cmbCSTORIG.BackColor = &HC0FFFF
End Sub

Private Sub cmbCSTORIG_Click()
   cmbPIS.SetFocus
End Sub

Private Sub CMBCSTORIG_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbPIS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CMBCSTORIG_KeyPress"
End Sub

Private Sub cmbCSTORIG_LostFocus()
   cmbCSTORIG.BackColor = &HFFFFFF
End Sub

Private Sub cmbPIS_GotFocus()
   cmbPIS.SelStart = 0
   cmbPIS.SelLength = Len(cmbPIS.Text)
   cmbPIS.BackColor = &HC0FFFF
End Sub

Private Sub cmbUFDestino_GotFocus()
   cmbUFDestino.SelStart = 0
   cmbUFDestino.SelLength = Len(cmbUFDestino.Text)
   cmbUFDestino.BackColor = &HC0FFFF
End Sub

Private Sub cmdLimpaBalanca_Click()
   chkPanific.Value = 0
   txtCasaInicioCodgProdBarra.Text = ""
   txtTamanhoCodgProdBarra.Text = ""
   txtTamanhoPesoValorBarra.Text = ""
   optGramas.Value = False
   optValor.Value = False
   txtCasaInicioCodgProdBarra.SetFocus
End Sub

Private Sub cmdSairBalanca_Click()
   Unload Me
End Sub

Private Sub chkContabiliza_Click()
   txtDiasPrazo.SetFocus
End Sub

Private Sub chkPermiteParcelar_Click()
   txtDiasPrazo.SetFocus
End Sub

Private Sub chkReceber_Click()
   txtDiasPrazo.SetFocus
End Sub

Private Sub chkPagar_Click()
   txtDiasPrazo.SetFocus
End Sub

Private Sub chkPermite_Desconto_Click()
   txtDiasPrazo.SetFocus
End Sub

Private Sub chkPreFatura_Desconto_Click()
   txtDiasPrazo.SetFocus
End Sub

Private Sub chkProducao_Click()
   txtDescUnidade.SetFocus
End Sub

Private Sub LISTADESCR_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTADESCR, ColumnHeader
End Sub

Private Sub LISTAPAGTO_Click()
On Error Resume Next

   txtFORMAPAGTO.Text = LISTAPAGTO.SelectedItem.Text
   txtFORMAPAGTO.SetFocus
End Sub

Private Sub LISTAPAGTO_DblClick()
On Error Resume Next

   txtFORMAPAGTO.Text = LISTAPAGTO.SelectedItem.Text
   txtFORMAPAGTO.SetFocus
End Sub

Private Sub LISTAVENDA_Click()
On Error Resume Next

   txtTIPOVENDA.Text = LISTAVENDA.SelectedItem.Text
End Sub

Private Sub LISTAVENDA_DblClick()
On Error Resume Next

   txtTIPOVENDA.Text = LISTAVENDA.SelectedItem.Text
   txtTIPOVENDA.SetFocus
End Sub

Private Sub chkCTesoraria_Click()
   txtDESCFORMAPAGTO.SetFocus
End Sub

Private Sub chkCBalcao_Click()
   txtDESCFORMAPAGTO.SetFocus
End Sub

Private Sub chkBaixaAuto_Click()
   txtDESCFORMAPAGTO.SetFocus
End Sub

Private Sub chkFunc_Click()
   txtDESCFORMAPAGTO.SetFocus
End Sub

Private Sub cmbdvsd_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbdvsf.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbdvsd_KeyPress"
End Sub

Private Sub cmbdvsf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbdved.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbdvsf_KeyPress"
End Sub

Private Sub cmbdved_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbdvef.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbdved_KeyPress"
End Sub

Private Sub cmbdvef_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbtrsd.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbdvef_KeyPress"
End Sub
Private Sub cmbtrsd_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbtrsf.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbtrsd_KeyPress"
End Sub

Private Sub cmbtrsf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbtred.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbtrsf_KeyPress"
End Sub

Private Sub cmbtred_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbtref.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbtred_KeyPress"
End Sub

Private Sub cmbtref_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtdecont.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbtref_KeyPress"
End Sub

Private Sub txtCFOP_ID_GotFocus()
   txtCFOP_ID.SelStart = 0
   txtCFOP_ID.SelLength = Len(txtCFOP_ID.Text)
   txtCFOP_ID.BackColor = &HC0FFFF
End Sub

Private Sub txtCFOP_ID_LostFocus()
   txtCFOP_ID.BackColor = &HFFFFFF
End Sub

Private Sub txtCOFINS_GotFocus()
   txtCOFINS.SelStart = 0
   txtCOFINS.SelLength = Len(txtCOFINS.Text)
   txtCOFINS.BackColor = &HC0FFFF
End Sub

Private Sub txtCOFINS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtSUBST.SetFocus
   End If
End Sub

Private Sub txtCOFINS_LostFocus()
   txtCOFINS.BackColor = &HFFFFFF
End Sub

Private Sub txtdecont_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtdencont.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdecont_KeyPress"
End Sub

Private Sub txtdencont_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtdemaq.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdencont_KeyPress"
End Sub
Private Sub txtdemaq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtdenmaq.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdemaq_KeyPress"
End Sub
Private Sub txtdenmaq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtfemaq.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdenmaq_KeyPress"
End Sub

Private Sub txtDesc_CFOP_GotFocus()
   txtDesc_CFOP.SelStart = 0
   txtDesc_CFOP.SelLength = Len(txtDesc_CFOP.Text)
   txtDesc_CFOP.BackColor = &HC0FFFF
End Sub

Private Sub txtDesc_CFOP_LostFocus()
   txtDesc_CFOP.BackColor = &HFFFFFF
End Sub

Private Sub cmbCOFINS_LostFocus()
   cmbCOFINS.BackColor = &HFFFFFF
End Sub

Private Sub cmbPIS_LostFocus()
   cmbPIS.BackColor = &HFFFFFF
End Sub

Private Sub txtfemaq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtfenmaq.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtfemaq_KeyPress"
End Sub
Private Sub txtfenmaq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtfeapa.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtfenmaq_KeyPress"
End Sub
Private Sub txtfeapa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     txtfenapa.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtfeapa_KeyPress"
End Sub
'============================= TAB000
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCodigo.Text) = "" Then
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select max(codigo) as ultimo_codg from DESCR WITH (NOLOCK)"
         SQL = SQL & " where TIPO = '" & Trim(txtTIPO.Text) & "'"
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            If Not IsNull(TabDESCR!ultimo_codg) Then
               txtCodigo.Text = TabDESCR!ultimo_codg + 1
               Else: txtCodigo.Text = 1
            End If
            Else: txtCodigo.Text = 1
         End If
      End If

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select * from descr WITH (NOLOCK)"
      SQL = SQL & " where TIPO = '" & Trim(txtTIPO.Text) & "'"
      SQL = SQL & " and codigo = '" & Trim(txtCodigo.Text) & "'"
      SQL = SQL & " order by codigo"
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         MOSTRA_DESC
      txtDesc.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcodigo_KeyPress"
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then _
      txtTIPO.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodigo_KeyUp"
End Sub

Private Sub txtCodigo_LostFocus()
'On Error GoTo ERRO_TRATA

   UCase (txtCodigo.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcodigo_LostFocus"
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   'KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select * from descr WITH (NOLOCK)"
      SQL = SQL & " where TIPO = '" & Trim(txtTIPO.Text) & "'"
      SQL = SQL & " and codigo = '" & Trim(txtCodigo.Text) & "'"
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         txtDesc.Text = UCase(txtDesc.Text)

      GRAVA_DESC
      LIMPA_DESC
      SETA_GRID_DESCR

      txtTIPO.Text = CRITERIO_A
      txtTIPO.SetFocus

      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesc_KeyPress"
End Sub

Private Sub cmdSAIRDESCR_Click()
'On Error GoTo ERRO_TRATA

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdSAIRdescr_Click"
End Sub

Private Sub cmdLIMPARdescr_Click()
'On Error GoTo ERRO_TRATA

   LIMPA_DESCR
   txtTIPO.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdLIMPARdescr_Click"
End Sub

Private Sub cmdMATARDESCR_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtTIPO.Text) <> "" And txtCodigo.Text <> "" Then
      SQL = "Delete from DESCR "
      SQL = SQL & " where TIPO = '" & Trim(txtTIPO.Text) & "'"
      SQL = SQL & " and codigo = '" & Trim(txtCodigo.Text) & "'"
      CONECTA_RETAGUARDA.Execute SQL

      txtTIPO.SetFocus
      SETA_GRID_DESCR
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdMATARDESCR_Click"
End Sub

Private Sub txtDesc_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   If KeyCode = 38 Then txtCodigo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDesc_KeyUp"
End Sub

Private Sub txtFISCO_GotFocus()
   txtFISCO.SelStart = 0
   txtFISCO.SelLength = Len(txtFISCO.Text)
   txtFISCO.BackColor = &HC0FFFF
End Sub

Private Sub txtFISCO_LostFocus()
   txtFISCO.BackColor = &HFFFFFF
End Sub

Private Sub txtICMS_Dentro_GotFocus()
   txtICMS_Dentro.SelStart = 0
   txtICMS_Dentro.SelLength = Len(txtICMS_Dentro.Text)
   txtICMS_Dentro.BackColor = &HC0FFFF
End Sub

Private Sub txtICMS_Dentro_LostFocus()
   txtICMS_Dentro.BackColor = &HFFFFFF
End Sub

Private Sub txtICMS_Fora_GotFocus()
   txtICMS_Fora.SelStart = 0
   txtICMS_Fora.SelLength = Len(txtICMS_Fora.Text)
   txtICMS_Fora.BackColor = &HC0FFFF
End Sub

Private Sub txtICMS_Fora_LostFocus()
   txtICMS_Fora.BackColor = &HFFFFFF
End Sub

Private Sub txtPIS_GotFocus()
   txtPIS.SelStart = 0
   txtPIS.SelLength = Len(txtPIS.Text)
   txtPIS.BackColor = &HC0FFFF
End Sub

Private Sub txtPIS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCOFINS.SetFocus
   End If
End Sub

Private Sub cmbCOFINS_Click()
   txtCOFINS.SetFocus
End Sub

Private Sub cmbCOFINS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCOFINS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCOFINS_KeyPress"
End Sub

Private Sub txtPIS_LostFocus()
   txtPIS.BackColor = &HFFFFFF
End Sub

Private Sub txtSUBST_GotFocus()
   txtSUBST.SelStart = 0
   txtSUBST.SelLength = Len(txtSUBST.Text)
   txtSUBST.BackColor = &HC0FFFF
End Sub

Private Sub txtSUBST_LostFocus()
   txtSUBST.BackColor = &HFFFFFF
End Sub

Private Sub txtBaseReduz_GotFocus()
   txtBaseReduz.SelStart = 0
   txtBaseReduz.SelLength = Len(txtBaseReduz.Text)
   txtBaseReduz.BackColor = &HC0FFFF
End Sub

Private Sub txtBaseReduz_LostFocus()
   txtBaseReduz.BackColor = &HFFFFFF
End Sub


Private Sub txtTipo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtDesc.Text = "" Then
         txtCodigo.SetFocus
         Exit Sub
      End If
      txtCodigo.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTipo_KeyPress"
End Sub

Private Sub LIMPA_DESCR()
'On Error GoTo ERRO_TRATA

   txtTIPO.Text = ""
   txtDesc.Text = ""
   txtCodigo.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_DESCR"
End Sub
'=========================== TAB6
Private Sub MOSTRA_PERC_CONT()
'On Error GoTo ERRO_TRATA

   Dim Cfop_Desc As String

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select empresa.* from EMPRESA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      'Entrada e Saida de Mercadoria
      If Not IsNull(TabTemp!CFOP_SAIDA_DE) Then
         cmbcfopsd.Text = (TabTemp!CFOP_SAIDA_DE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_SAIDA_DE)
      End If
      If Not IsNull(TabTemp!CFOP_SAIDA_FE) Then
         cmbcfopsf.Text = (TabTemp!CFOP_SAIDA_FE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_SAIDA_FE)
      End If
      If Not IsNull(TabTemp!CFOP_ENTRADA_DE) Then
         cmbcfoped.Text = (TabTemp!CFOP_ENTRADA_DE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_ENTRADA_DE)
      End If
      If Not IsNull(TabTemp!CFOP_ENTRADA_FE) Then
         cmbcfopef.Text = (TabTemp!CFOP_ENTRADA_FE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_ENTRADA_FE)
      End If
      'Devolução de Entrada e Saida
      If Not IsNull(TabTemp!CFOP_DV_SAI_DE) Then
         cmbdvsd.Text = (TabTemp!CFOP_DV_SAI_DE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_DV_SAI_DE)
      End If
      If Not IsNull(TabTemp!CFOP_DV_SAI_FE) Then
         cmbdvsf.Text = (TabTemp!CFOP_DV_SAI_FE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_DV_SAI_FE)
      End If
      If Not IsNull(TabTemp!CFOP_DV_ENT_DE) Then
         cmbdved.Text = (TabTemp!CFOP_DV_ENT_DE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_DV_ENT_DE)
      End If
      If Not IsNull(TabTemp!CFOP_DV_ENT_FE) Then
         cmbdvef.Text = (TabTemp!CFOP_DV_ENT_FE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_DV_ENT_FE)
      End If
      'Transferencia Entrada e Saida
      If Not IsNull(TabTemp!CFOP_TRA_SAI_DE) Then
         cmbtrsd.Text = (TabTemp!CFOP_TRA_SAI_DE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_TRA_SAI_DE)
      End If
      If Not IsNull(TabTemp!CFOP_TRA_SAI_FE) Then
         cmbtrsf.Text = (TabTemp!CFOP_TRA_SAI_FE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_TRA_SAI_FE)
      End If
      If Not IsNull(TabTemp!CFOP_TRA_ENT_DE) Then
         cmbtred.Text = (TabTemp!CFOP_TRA_ENT_DE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_TRA_ENT_DE)
      End If
      If Not IsNull(TabTemp!CFOP_TRA_ENT_FE) Then
         cmbtref.Text = (TabTemp!CFOP_TRA_ENT_FE) & "-" & RetornaDescricaoCFOP(TabTemp!CFOP_TRA_ENT_FE)
      End If
      
      'Cadastro de Percentuais de reducao de Base de Calculo
      If Not IsNull(TabTemp!TP2_DE_CONTRIB) Then
         txtdecont.Text = TabTemp!TP2_DE_CONTRIB
      End If
      If Not IsNull(TabTemp!TP2_DE_NCONTRIB) Then
         txtdencont.Text = TabTemp!TP2_DE_NCONTRIB
      End If
      If Not IsNull(TabTemp!TP2_DE_CMAQ_IMP) Then
         txtdemaq.Text = TabTemp!TP2_DE_CMAQ_IMP
      End If
      If Not IsNull(TabTemp!TP2_DE_NMAQ_IMP) Then
         txtdenmaq.Text = TabTemp!TP2_DE_NMAQ_IMP
      End If
      If Not IsNull(TabTemp!TP2_FE_CMAQ_IMP) Then
         txtfemaq.Text = TabTemp!TP2_FE_CMAQ_IMP
      End If
      If Not IsNull(TabTemp!TP2_FE_NMAQ_IMP) Then
         txtfenmaq.Text = TabTemp!TP2_FE_NMAQ_IMP
      End If
      If Not IsNull(TabTemp!TP2_FE_CAP_INDU) Then
         txtfeapa.Text = TabTemp!TP2_FE_CAP_INDU
      End If
      If Not IsNull(TabTemp!TP2_FE_CAP_INDU) Then
         txtfenapa.Text = TabTemp!TP2_FE_CAP_INDU
      End If
      'fazer assim para as outras
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PERC_CONT"
End Sub

Private Sub GRAVA_PERC_CONT()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select empresa.* from EMPRESA WITH (NOLOCK)"

   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      'entrada e saida
      SQL = "UPDATE EMPRESA SET "
      SQL = SQL & " CFOP_SAIDA_DE = " & Left(cmbcfopsd.Text, 4)
      SQL = SQL & ", CFOP_SAIDA_FE = " & Left(cmbcfopsf.Text, 4)
      SQL = SQL & ", CFOP_ENTRADA_DE = " & Left(cmbcfoped.Text, 4)
      SQL = SQL & ", CFOP_ENTRADA_FE = " & Left(cmbcfopef.Text, 4)
      SQL = SQL & ","
      'devolucoes
      SQL = SQL & " CFOP_DV_SAI_DE = " & Left(cmbdvsd.Text, 4) & ", CFOP_DV_SAI_FE  = " & Left(cmbdvsf.Text, 4) & ","
      SQL = SQL & " CFOP_DV_ENT_DE  = " & Left(cmbdved.Text, 4) & ", CFOP_DV_ENT_FE  = " & Left(cmbdvef.Text, 4) & ","
      'Transferencias
      SQL = SQL & " CFOP_TRA_SAI_DE  = " & Left(cmbtrsd.Text, 4) & ", CFOP_TRA_SAI_FE  = " & Left(cmbtrsf.Text, 4) & ","
      SQL = SQL & " CFOP_TRA_ENT_DE  = " & Left(cmbtred.Text, 4) & ", CFOP_TRA_ENT_FE = " & Left(cmbtref.Text, 4) & ","
      'impostos
      SQL = SQL & " TP2_DE_CONTRIB = " & txtdecont.Text & ", TP2_DE_NCONTRIB  = " & txtdencont.Text & ","
      SQL = SQL & " TP2_DE_CMAQ_IMP  = " & txtdemaq.Text & ", TP2_DE_NMAQ_IMP = " & txtdenmaq.Text & ","
      SQL = SQL & " TP2_FE_CMAQ_IMP  = " & txtfemaq.Text & ", TP2_FE_NMAQ_IMP = " & txtfenmaq.Text & ","
      SQL = SQL & " TP2_FE_CAP_INDU  = " & txtfeapa.Text & ", TP2_FE_NAP_INDU = " & txtfenapa.Text & ","

      SQL = SQL & " from EMPRESA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
      SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
      SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      CONECTA_RETAGUARDA.Execute SqL2
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtdecont.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_PERC_CONT"
End Sub
'==============================tab01
Private Sub txttipovenda_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe o código", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txttipovenda_GotFocus"
End Sub

Private Sub txtTIPOVENDA_keypress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtTIPOVENDA.Text) = "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select max(TIPOVENDA_id) from TIPOVENDA WITH (NOLOCK)"
         SQL = SQL & " where TIPOVENDA_id < 9999"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            If Not IsNull(TabTemp.Fields(0).Value) Then
               txtTIPOVENDA.Text = TabTemp.Fields(0).Value + 1
               Else: txtTIPOVENDA.Text = 1
            End If
            Else: txtTIPOVENDA.Text = 1
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      chkPermite_Desconto.Value = 0
      chkPreFatura.Value = 0
      chkPagar.Value = 0
      chkReceber.Value = 0
      chkContabiliza.Value = 0
      chkPermiteParcelar.Value = 0

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " where tipovenda_id = " & txtTIPOVENDA.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
            If TabTemp.Fields("permite_desconto").Value = True Then
               chkPermite_Desconto.Value = 1
               Else: chkPermite_Desconto.Value = 0
            End If
         End If
         If Not IsNull(TabTemp.Fields("PreFatura").Value) Then
            If TabTemp.Fields("PreFatura").Value = True Then
               chkPreFatura.Value = 1
               Else: chkPreFatura.Value = 0
            End If
         End If
         If Not IsNull(TabTemp.Fields("PAGAR").Value) Then
            If TabTemp.Fields("PAGAR").Value = True Then
               chkPagar.Value = 1
               Else: chkPagar.Value = 0
            End If
         End If
         If Not IsNull(TabTemp.Fields("RECEBER").Value) Then
            If TabTemp.Fields("RECEBER").Value = True Then
               chkReceber.Value = 1
               Else: chkReceber.Value = 0
            End If
         End If
         If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
            If TabTemp.Fields("contabiliza").Value = True Then
               chkContabiliza.Value = 1
               Else: chkContabiliza.Value = 0
            End If
         End If
         If Not IsNull(TabTemp.Fields("PERMITEPARCELAR").Value) Then
            If TabTemp.Fields("PERMITEPARCELAR").Value = True Then
               chkPermiteParcelar.Value = 1
               Else: chkPermiteParcelar.Value = 0
            End If
         End If
         If Not IsNull(TabTemp.Fields("cartaoadm_id").Value) Then
            If TabTemp.Fields("cartaoadm_id").Value > 0 Then
               cmbADMCartaoAUX.Text = "" & TabTemp.Fields("cartaoadm_id").Value

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close
               SQL = "select * from CARTAOADM WITH (NOLOCK)"
               SQL = SQL & " where cartaoadm_id = " & cmbADMCartaoAUX.Text
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabConsulta.EOF Then _
                  cmbADMCARTAO.Text = "" & Trim(TabConsulta.Fields("fantasia").Value)
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close
            End If
         End If

         txtDESCTIPOVENDA.Text = "" & TabTemp!DESCRICAO
         txtParcela.Text = "" & TabTemp!parcela
         txtDiasPrazo.Text = "" & TabTemp!PRAZO
         txtDiaVencto.Text = "" & TabTemp.Fields("DIAVENCTO").Value

         txtPercJuros.Text = "" & TabTemp!PERC_JUROS
         cmbAuxForma.Text = "" & TabTemp!FORMAPAGTO_ID
         txtDebito.Text = "" & TabTemp.Fields("PERC_CARTAO_DEBITO").Value
         txtCredito.Text = "" & TabTemp.Fields("PERC_CARTAO_credito").Value

         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
         SQL = SQL & " where formapagto_id = " & cmbAuxForma.Text
         SQL = SQL & " and status = 'true' "
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            cmbForma.Text = TabAUX!DESCRICAO
         If TabAUX.State = 1 Then _
            TabAUX.Close

         If Not IsNull(TabTemp.Fields("CC_ID").Value) Then
            cmbCCAux.Text = TabTemp.Fields("CC_ID").Value

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select * from TIPOVENDA c WITH (NOLOCK)"
            SQL = SQL & " left join DESCR d WITH (NOLOCK)"
            SQL = SQL & " on c.cc_id = d.codigo "
            SQL = SQL & " where d.TIPO = 'O' "
            SQL = SQL & " and d.codigo = " & cmbCCAux.Text
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
               cmbCC.Text = "" & TabTemp.Fields("descricao").Value
            If TabTemp.State = 1 Then _
               TabTemp.Close
            Else
               cmbCCAux.Text = ""
               cmbCC.Text = ""
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      txtDESCTIPOVENDA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txttipovenda_KeyPress"
End Sub

Private Sub txtdesctipovenda_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbForma.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtdesctipovenda_KeyPress"
End Sub

Private Sub txtparcela_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDiasPrazo.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtparcela_KeyPress"
End Sub

Private Sub cmbforma_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtParcela.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbforma_KeyPress"
End Sub

Private Sub txtpercjuros_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If UCase(Trim(Mid(cmbForma.Text, 1, 6))) = UCase("cartão") Or UCase(Trim(Mid(cmbForma.Text, 1, 6))) = UCase("cartAo") Then
         txtDebito.SetFocus
         Exit Sub
         Else: PROCESSA_REGISTRO
      End If

      txtTIPOVENDA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtpercjuros_KeyPress"
End Sub

Private Sub cmbFORMA_Click()
'On Error GoTo ERRO_TRATA

   cmbAuxForma.ListIndex = cmbForma.ListIndex
   txtParcela.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFORMA_Click"
End Sub

Private Sub txtDIASprazo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDiaVencto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDIASprazo_KeyPress"
End Sub

Private Sub txtDiaVencto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercJuros.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDiaVencto_KeyPress"
End Sub

Private Sub txtDebito_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtCredito.SetFocus

      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDebito_KeyPress"
End Sub

Private Sub txtCredito_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      PROCESSA_REGISTRO
      txtTIPOVENDA.SetFocus

      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCredito_KeyPress"
End Sub

Private Sub cmbCC_Click()
On Error Resume Next

   cmbCCAux.ListIndex = cmbCC.ListIndex
   txtPercJuros.SetFocus
End Sub

Private Sub cmbADMCARTAO_Click()
On Error Resume Next

   cmbADMCartaoAUX.ListIndex = cmbADMCARTAO.ListIndex
   txtPercJuros.SetFocus
End Sub

Private Sub cmdSAIRvenda_Click()
'On Error GoTo ERRO_TRATA

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdSAIRvenda_Click"
End Sub

Private Sub cmdLIMPARvenda_Click()
'On Error GoTo ERRO_TRATA

   LIMPA_TIPOVENDA
   txtTIPOVENDA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdLIMPARvenda_Click"
End Sub

Private Sub cmdmatarvenda_Click()
'On Error GoTo ERRO_TRATA
   If txtTIPOVENDA.Text <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " where tipovenda_id = " & txtTIPOVENDA.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Confirma exclusão ?"
         PERGUNTA Msg, vbYesNo + 32, "Devolução NFE", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            CONECTA_RETAGUARDA.Execute "Delete from TIPOVENDA where tipovenda_id = " & txtTIPOVENDA.Text
            LIMPA_TIPOVENDA
            SETA_GRID_VENDA
            txtTIPOVENDA.SetFocus
            Exit Sub
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   txtTIPOVENDA.SetFocus
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdmatarvenda_Click"
End Sub

Private Sub SETA_GRID_VENDA()
'On Error GoTo ERRO_TRATA

   LISTAVENDA.ListItems.Clear
   CONT_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      CONT_N = CONT_N + 1
      Set item = LISTAVENDA.ListItems.Add(, "seq." & CONT_N, TabTemp!TIPOVENDA_ID)
      item.SubItems(1) = "" & TabTemp!DESCRICAO
      item.SubItems(2) = "" & TabTemp!parcela

      If Not IsNull(TabTemp!FORMAPAGTO_ID) Then
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select descricao from FORMAPAGTO WITH (NOLOCK)"
         SQL = SQL & " where formapagto_id = " & TabTemp!FORMAPAGTO_ID
         SQL = SQL & " and status = 'true' "
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            item.SubItems(3) = "" & TabDESCR!DESCRICAO
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
      End If
      If Not IsNull(TabTemp!PRAZO) Then _
         item.SubItems(4) = TabTemp!PRAZO & " dias"

      If Not IsNull(TabTemp!PERC_JUROS) Then _
         item.SubItems(5) = TabTemp!PERC_JUROS & " %"

      If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
         If TabTemp.Fields("contabiliza").Value = True Then
            item.SubItems(6) = "Sim"
            Else: item.SubItems(6) = "Não"
         End If
      End If
      If Not IsNull(TabTemp.Fields("permite_desconto").Value) Then
         If TabTemp.Fields("permite_desconto").Value = True Then
            item.SubItems(7) = "Sim"
            Else: item.SubItems(7) = "Não"
         End If
      End If
      If Not IsNull(TabTemp.Fields("CC_ID").Value) Then
         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from DESCR d WITH (NOLOCK)"
         SQL = SQL & " where TIPO = 'O' "
         SQL = SQL & " and codigo = '" & Trim(TabTemp.Fields("CC_ID").Value) & "'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then _
            item.SubItems(8) = "" & TabAUX.Fields("descricao").Value
         If TabAUX.State = 1 Then _
            TabAUX.Close
      End If
      If Not IsNull(TabTemp.Fields("pagar").Value) Then
         If TabTemp.Fields("pagar").Value = True Then
            item.SubItems(10) = "Sim"
            Else: item.SubItems(10) = "Não"
         End If
      End If
      If Not IsNull(TabTemp.Fields("receber").Value) Then
         If TabTemp.Fields("receber").Value = True Then
            item.SubItems(11) = "Sim"
            Else: item.SubItems(11) = "Não"
         End If
      End If
      If Not IsNull(TabTemp.Fields("permiteparcelar").Value) Then
         If TabTemp.Fields("permiteparcelar").Value = True Then
            item.SubItems(12) = "Sim"
            Else: item.SubItems(12) = "Não"
         End If
      End If
      item.SubItems(13) = ""
      If Not IsNull(TabTemp.Fields("cartaoadm_id").Value) Then
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "select fantasia from CARTAOADM WITH (NOLOCK)"
         SQL = SQL & " where cartaoadm_id = " & TabTemp.Fields("cartaoadm_id").Value
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            item.SubItems(13) = "" & Trim(TabConsulta.Fields("fantasia").Value)
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_VENDA"
End Sub

Private Sub LIMPA_TIPOVENDA()
'On Error GoTo ERRO_TRATA

   chkPagar.Value = 0
   chkReceber.Value = 0
   chkContabiliza.Value = 0
   chkPermite_Desconto.Value = 0
   chkPreFatura.Value = 0
   txtDebito.Text = ""
   txtCredito.Text = ""
   txtPercJuros.Text = ""
   txtTIPOVENDA.Text = ""
   txtDESCTIPOVENDA.Text = ""
   txtParcela.Text = ""
   txtDiasPrazo.Text = ""
   txtDiaVencto.Text = ""
   cmbAuxForma.Text = ""
   cmbForma.Text = ""
   cmbCCAux.Text = ""
   cmbCC.Text = ""
   cmbADMCARTAO.Text = ""
   cmbADMCartaoAUX.Text = ""

   SETA_GRID_VENDA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TIPOVENDA"
End Sub
'===================================
'=================SUBROTINAS
Private Sub GRAVA_DESC()
'On Error GoTo ERRO_TRATA

   If TabDESCR.EOF Then
      SqL2 = "INSERT INTO DESCR (TIPO, Codigo, DESCRICAO) "
      SqL2 = SqL2 & " VALUES ('" & Trim(txtTIPO.Text) & "','" & Trim(txtCodigo.Text) & "','" & txtDesc.Text & "')"
      CONECTA_RETAGUARDA.Execute SqL2
      Else: CONECTA_RETAGUARDA.Execute "UPDATE DESCR SET TIPO = '" & Trim(txtTIPO.Text) & "', Codigo = '" & Trim(txtCodigo.Text) & "', DESCRICAO = '" & txtDesc.Text & "' where TIPO = '" & Trim(txtTIPO.Text) & "' and codigo = " & txtCodigo.Text
   End If
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_DESC"
End Sub

Private Sub LIMPA_DESC()
'On Error GoTo ERRO_TRATA
   txtTIPO.Text = ""
   txtCodigo.Text = ""
   txtDesc.Text = ""
   SETA_GRID_DESCR
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_DESC"
End Sub

Private Sub MOSTRA_DESC()
'On Error GoTo ERRO_TRATA
   txtDesc.Text = Trim(TabDESCR!DESCRICAO)
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DESC"
End Sub

Private Sub List1_Click()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Left(List1.Text, 2)
   SETA_GRID_DESCR
   txtTIPO.Text = CRITERIO_A
   txtTIPO.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "List1_Click"
End Sub

Private Sub SETA_GRID_DESCR()
'On Error GoTo ERRO_TRATA

   LISTADESCR.ListItems.Clear
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from DESCR WITH (NOLOCK)"
   SQL = SQL & " where TIPO = '" & CRITERIO_A & "'"
   SQL = SQL & " order by codigo"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = LISTADESCR.ListItems.Add(, "seq." & TabTemp!Codigo, Trim(TabTemp.Fields("codigo").Value))
      item.SubItems(1) = TabTemp!TIPO
      item.SubItems(2) = Trim(TabTemp!DESCRICAO)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_DESCR"
End Sub

Private Sub Mata_Desc()
'On Error GoTo ERRO_TRATA
   If Trim(txtTIPO.Text) <> "" And txtCodigo.Text <> "" Then
      SQL = "select * from descr WITH (NOLOCK)"
      SQL = SQL & " where TIPO = '" & Trim(txtTIPO.Text) & "'"
      SQL = SQL & " and codigo = '" & Trim(txtCodigo.Text) & "'"
      SQL = SQL & " order by codigo"
      If TabDESCR.State = 1 Then TabDESCR.Close
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then
         SQL = "Delete from DESCR "
         SQL = SQL & " where TIPO = '" & Trim(txtTIPO.Text) & "'"
         SQL = SQL & " and codigo = '" & Trim(txtCodigo.Text) & "'"
         CONECTA_RETAGUARDA.Execute SQL

         LIMPA_DESC
         'dataDesc.Refresh
         txtTIPO.SetFocus
         'dataDesc.Refresh
      End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Mata_Desc"
End Sub

'==============================tab01
Private Sub txtFormaPagto_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe o código", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtFormaPagto_GotFocus"
End Sub

Private Sub TXTFORMAPAGTO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtFORMAPAGTO.Text) = "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select max(formapagto_id) from FORMAPAGTO WITH (NOLOCK)"
         SQL = SQL & " where formapagto_id < 9999"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            If Not IsNull(TabTemp.Fields(0).Value) Then
               txtFORMAPAGTO.Text = TabTemp.Fields(0).Value + 1
               Else: txtFORMAPAGTO.Text = 1
            End If
            Else: txtFORMAPAGTO.Text = 1
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where formapagto_id = " & txtFORMAPAGTO.Text
      SQL = SQL & " and status = 'true' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("contab_tesora").Value) Then
            If TabTemp.Fields("contab_tesora").Value = True Then
               chkCTesoraria.Value = 1
               Else: chkCTesoraria.Value = 0
            End If
         End If

         If Not IsNull(TabTemp.Fields("contab_balcao").Value) Then
            If TabTemp.Fields("contab_balcao").Value = True Then
               chkCBalcao.Value = 1
               Else: chkCBalcao.Value = 0
            End If
         End If

         If Not IsNull(TabTemp.Fields("baixaauto").Value) Then
            If TabTemp.Fields("baixaauto").Value = True Then
               chkBaixaAuto.Value = 1
               Else: chkBaixaAuto.Value = 0
            End If
         End If

         If Not IsNull(TabTemp.Fields("FUNC").Value) Then
            If TabTemp.Fields("FUNC").Value = True Then
               chkFunc.Value = 1
               Else: chkFunc.Value = 0
            End If
         End If

         If Not IsNull(TabTemp!DESCRICAO) Then _
            txtDESCFORMAPAGTO.Text = Trim(TabTemp!DESCRICAO)

         If Not IsNull(TabTemp.Fields("status").Value) Then
            If TabTemp.Fields("status").Value = True Then
               chkPagto.Value = 1
               Else: chkPagto.Value = 0
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      txtDESCFORMAPAGTO.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTFORMAPAGTO_KeyPress"
End Sub

Private Sub TXTDESCFORMAPAGTO_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe a descrição", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTDESCFORMAPAGTO_GotFocus"
End Sub

Private Sub txtDescFormaPagto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txtFORMAPAGTO.Text) = "" Then
         MsgBox "Código inválido."
         txtFORMAPAGTO.SetFocus
         Exit Sub
      End If
      If Trim(txtDESCFORMAPAGTO.Text) = "" Then
         MsgBox "Descrição inválida."
         txtDESCFORMAPAGTO.SetFocus
         Exit Sub
      End If
      KeyAscii = 0

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where formapagto_id = " & txtFORMAPAGTO.Text
      SQL = SQL & " and status = 'true' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "INSERT INTO FORMAPAGTO "
         SQL = SQL & " ("
            SQL = SQL & " empresa_id, formapagto_id, Descricao, "
            SQL = SQL & " Status, contab_tesora,"
            SQL = SQL & " BAIXAAUTO,contab_balcao,Func"
         SQL = SQL & " ) "
         SQL = SQL & " VALUES ("
            SQL = SQL & EMPRESA_ID_N                           'empresa_id
            SQL = SQL & "," & txtFORMAPAGTO.Text               'formapagto_id
            SQL = SQL & ",'" & txtDESCFORMAPAGTO.Text & "'"    'Descricao
            SQL = SQL & "," & chkPagto.Value                   'Status
            SQL = SQL & "," & chkCTesoraria.Value              'contab_tesora
            SQL = SQL & "," & chkBaixaAuto.Value               'baixa auto
            SQL = SQL & "," & chkCBalcao.Value                 'contab_BALCAO
            SQL = SQL & "," & chkFunc.Value                    'FUNC
         SQL = SQL & ")"
         Else
            SQL = "UPDATE FORMAPAGTO SET "
            SQL = SQL & " formapagto_id = " & Trim(txtFORMAPAGTO.Text)
            SQL = SQL & ", Descricao = '" & Trim(txtDESCFORMAPAGTO.Text) & "'"
            SQL = SQL & ", Status = " & chkPagto.Value                   'Status
            SQL = SQL & ", contab_tesora = " & chkCTesoraria.Value
            SQL = SQL & ", contab_BALCAO = " & chkCBalcao.Value
            SQL = SQL & ", baixaauto = " & chkBaixaAuto.Value
            SQL = SQL & ", FUNC = " & chkFunc.Value                    'FUNC
            SQL = SQL & " where formapagto_id = " & txtFORMAPAGTO.Text
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      CONECTA_RETAGUARDA.Execute SQL

      SETA_GRID_PAGTO
      LIMPA_FORMAPAGTO
      txtFORMAPAGTO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescFormaPagto_KeyPress"
End Sub

Private Sub cmdSAIRPAGTO_Click()
'On Error GoTo ERRO_TRATA
   Unload Me
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdSAIRPAGTO_Click"
End Sub

Private Sub cmdLIMPARPAGTO_Click()
'On Error GoTo ERRO_TRATA

   LIMPA_FORMAPAGTO
   txtFORMAPAGTO.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdLIMPARPAGTO_Click"
End Sub

Private Sub cmdmatarPAGTO_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtFORMAPAGTO.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where FORMAPAGTO_ID = " & txtFORMAPAGTO.Text
      'SQL = SQL & " and status = 'true' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         'ITEM DE LANÇAMENTO
         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from ITEMLANCAMENTO WITH (NOLOCK)"
         SQL = SQL & " where FORMAPAGTO_ID = " & TabTemp!FORMAPAGTO_ID
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            If TabAUX.State = 1 Then _
               TabAUX.Close

            MsgBox "Impossivel excluir, já existe registro relacionado. ITEMLANCAMENTO"
            Exit Sub
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         'ITEM DE CAIXA
         SQL = "select * from CAIXADIA WITH (NOLOCK)"
         SQL = SQL & " INNER JOIN CAIXADIAITEM WITH (NOLOCK)"
         SQL = SQL & " ON CAIXADIA.CAIXADIA_ID = CAIXADIAITEM.CAIXADIA_ID "

         SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and NUMERO_CAIXA_CPU = " & NUMERO_CAIXA_CPU

         SQL = SQL & " and CAIXADIAITEM.caixadia_id = " & CAIXA_DIA_ID_N
         SQL = SQL & " and usuario_id = " & USUARIO_ID_N
         SQL = SQL & " and formapagto_id = " & TabTemp!FORMAPAGTO_ID

         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            If TabAUX.State = 1 Then _
               TabAUX.Close

            MsgBox "Impossivel excluir, já existe registro relacionado. CAIXADIAITEM"
            Exit Sub
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         'ITEM DE CAIXATESORARIA
         SQL = "select * from CAIXATESORARIAITEM WITH (NOLOCK)"
         SQL = SQL & " where FORMAPAGTO_ID = " & TabTemp!FORMAPAGTO_ID
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            If TabAUX.State = 1 Then _
               TabAUX.Close

            MsgBox "Impossivel excluir, já existe registro relacionado. CAIXATESORARIAITEM"
            Exit Sub
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         Msg = "Confirma exclusão ?"
         PERGUNTA Msg, vbYesNo + 32, "Matar Pagamento", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            SQL = "Delete from CAIXATESORARIAITEM "
            SQL = SQL & " where FORMAPAGTO_ID = " & TabTemp!FORMAPAGTO_ID
            CONECTA_RETAGUARDA.Execute SQL

            SQL = "Delete from FORMAPAGTO"
            SQL = SQL & " where FORMAPAGTO_ID = " & TabTemp!FORMAPAGTO_ID
            CONECTA_RETAGUARDA.Execute SQL

            LIMPA_FORMAPAGTO
            SETA_GRID_PAGTO

            txtFORMAPAGTO.SetFocus
            Exit Sub
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   txtFORMAPAGTO.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdmatarPAGTO_Click"
End Sub

Private Sub LIMPA_FORMAPAGTO()
'On Error GoTo ERRO_TRATA

   chkCTesoraria.Value = 0
   chkCBalcao.Value = 0
   chkBaixaAuto.Value = 0
   chkFunc.Value = 0
   txtFORMAPAGTO.Text = ""
   txtDESCFORMAPAGTO.Text = ""
   SETA_GRID_PAGTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_FORMAPAGTO"
End Sub
'==============================
Private Sub TXTTIPOENTRADA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtTIPOENTRADA.Text = "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select max(TIPOENTRADA_id) as ultimo_codg from TIPOENTRADA WITH (NOLOCK)"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            If Not IsNull(TabTemp!ultimo_codg) Then
               txtTIPOENTRADA.Text = TabTemp!ultimo_codg + 1
               Else: txtTIPOENTRADA.Text = 1
            End If
            Else: txtTIPOENTRADA.Text = 1
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
         Else
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select descricao from TIPOENTRADA WITH (NOLOCK)"
            SQL = SQL & " where TIPOENTRADA_id = " & txtTIPOENTRADA.Text
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
               If Not IsNull(TabTemp!DESCRICAO) Then _
                  txtDESCTIPOENTRADA.Text = TabTemp!DESCRICAO
            If TabTemp.State = 1 Then _
               TabTemp.Close
      End If
      txtDESCTIPOENTRADA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTTIPOENTRADA_KeyPress"
End Sub

Private Sub TXTDESCTIPOENTRADA_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If txtTIPOENTRADA.Text = "" Then
         MsgBox "Código inválido."
         txtTIPOENTRADA.SetFocus
         Exit Sub
      End If
      If txtDESCTIPOENTRADA.Text = "" Then
         MsgBox "Descrição inválida."
         txtDESCTIPOENTRADA.SetFocus
         Exit Sub
      End If
      KeyAscii = 0

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TIPOENTRADA WITH (NOLOCK)"
      SQL = SQL & " where TIPOENTRADA_id = " & txtTIPOENTRADA.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SqL2 = "INSERT INTO TIPOENTRADA (TIPOENTRADA_id, Descricao) "
         SqL2 = SqL2 & " VALUES (" & txtTIPOENTRADA.Text & ",'" & txtDESCTIPOENTRADA.Text & "')"
         CONECTA_RETAGUARDA.Execute SqL2
      Else
         SqL2 = "UPDATE TIPOENTRADA SET TIPOENTRADA_id  = " & txtTIPOENTRADA.Text & ", Descricao = '" & txtDESCTIPOENTRADA.Text & "'"
         SqL2 = SqL2 & " where TIPOENTRADA_id = " & txtTIPOENTRADA.Text
         CONECTA_RETAGUARDA.Execute SqL2
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SETA_GRID_ENTRADA
      LIMPA_TIPOENTRADA
      txtTIPOENTRADA.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTDESCTIPOENTRADA_KeyPress"
End Sub

Private Sub cmdSAIRentrada_Click()
'On Error GoTo ERRO_TRATA

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdSAIRentrada_Click"
End Sub

Private Sub cmdLIMPARentrada_Click()
'On Error GoTo ERRO_TRATA

   LIMPA_TIPOENTRADA
   txtTIPOENTRADA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdLIMPARentrada_Click"
End Sub

Private Sub cmdmatarentrada_Click()
'On Error GoTo ERRO_TRATA

   If txtTIPOENTRADA.Text <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      
      SQL = "select * from TIPOENTRADA WITH (NOLOCK)"
      SQL = SQL & " where TIPOENTRADA_id = " & txtTIPOENTRADA.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Confirma exclusão ?"
         PERGUNTA Msg, vbYesNo + 32, "Matar Entrada", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            CONECTA_RETAGUARDA.Execute "delete from TIPOENTRADA where TIPOENTRADA_id = " & txtTIPOENTRADA.Text
            LIMPA_TIPOENTRADA
            SETA_GRID_VENDA
            txtTIPOENTRADA.SetFocus
            Exit Sub
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   txtTIPOENTRADA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdmatarentrada_Click"
End Sub

Private Sub SETA_GRID_ENTRADA()
'On Error GoTo ERRO_TRATA

   LISTAENTRADA.ListItems.Clear
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOENTRADA WITH (NOLOCK)"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = LISTAENTRADA.ListItems.Add(, "seq." & TabTemp!TIPOENTRADA_id, TabTemp!TIPOENTRADA_id)
      item.SubItems(1) = TabTemp!DESCRICAO
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_ENTRADA"
End Sub

Private Sub LIMPA_TIPOENTRADA()
'On Error GoTo ERRO_TRATA

   txtTIPOENTRADA.Text = ""
   txtDESCTIPOENTRADA.Text = ""
   SETA_GRID_ENTRADA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TIPOENTRADA"
End Sub

'===================================
Private Sub SETA_GRID_CFOPold()
'On Error GoTo ERRO_TRATA

   lstCFOP.ListItems.Clear
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CFOP WITH (NOLOCK)"
   SQL = SQL & "order by cfop_id "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstCFOP.ListItems.Add(, "seq." & TabTemp!CFOP_ID, TabTemp!CFOP_ID)
      item.SubItems(1) = "" & Trim(TabTemp!DESCRICAO)
'      item.SubItems(3) = "" & Trim(TabTemp!ALIQUOTA_ICMS_DENTRO)
      item.SubItems(4) = "" & Trim(TabTemp!PERC_IPI)
      item.SubItems(5) = "" & Trim(TabTemp!PERC_ISS)
      item.SubItems(6) = "" & Trim(TabTemp!PERC_SUBST_TRIB)
      item.SubItems(2) = "" & Trim(TabTemp!MSGFISCO)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_CFOP"
End Sub

Private Sub SETA_GRID_CFOP()
'On Error GoTo ERRO_TRATA

   Dim UF_DESTINO_A As String
   UF_DESTINO_A = ""

   lstCFOP.ListItems.Clear
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT CFOP.CFOP_ID, CFOP.ESTABELECIMENTO_ID, CFOP.DESCRICAO, CFOP.PERC_SUBST_TRIB, "
   SQL = SQL & " CFOPUF.CFOPUF_ID, CFOPUF.UF_ORIGEM, CFOPUF.UF_DESTINO, ALIQUOTA_UF.CST_ICMS, "
   SQL = SQL & " ALIQUOTA_UF.ALIQUOTA_ICMS_DENTRO, ALIQUOTA_UF.ALIQUOTA_ICMS_FORA, ALIQUOTA_UF.CST_PIS, "
   SQL = SQL & " ALIQUOTA_UF.ALIQUOTA_PIS, ALIQUOTA_UF.CST_COFINS, ALIQUOTA_UF.ALIQUOTA_COFINS,perc_base_reduz"
   SQL = SQL & " FROM ALIQUOTA_UF WITH (NOLOCK)"
   SQL = SQL & "INNER JOIN CFOPUF WITH (NOLOCK)"
   SQL = SQL & "ON ALIQUOTA_UF.CFOPUF_ID = CFOPUF.CFOPUF_ID "
   SQL = SQL & "INNER JOIN CFOP WITH (NOLOCK)"
   SQL = SQL & "ON CFOPUF.CFOP_ID = CFOP.CFOP_ID"

SQL = SQL & " where uf_origem = '" & Trim(txtUFORIGEM.Text) & "'"

   SQL = SQL & " order by UF_origem,CFOP.CFOP_ID "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      UF_DESTINO_A = "" & Trim(Trim(TabTemp!UF_DESTINO))
      If UF_DESTINO_A = "" Then
         UF_DESTINO_A = "Todos"
      End If
      Set item = lstCFOP.ListItems.Add(, "seq." & TabTemp!CFOPUF_ID, TabTemp!CFOPUF_ID)

      item.SubItems(1) = "" & Trim(TabTemp!CFOP_ID)
      item.SubItems(2) = "" & Trim(TabTemp!DESCRICAO)
      item.SubItems(3) = "" & Trim(TabTemp!UF_origem)
      item.SubItems(4) = "" & TabTemp!CST_ICMS
      item.SubItems(5) = "" & Format(TabTemp!ALIQUOTA_ICMS_DENTRO, strFormatacao2Digitos)
      item.SubItems(6) = "" & UF_DESTINO_A
      item.SubItems(7) = "" & Format(TabTemp!ALIQUOTA_ICMS_FORA, strFormatacao2Digitos)
      item.SubItems(8) = "" & TabTemp!CST_PIS
      item.SubItems(9) = "" & Format(TabTemp!ALIQUOTA_PIS, strFormatacao2Digitos)
      item.SubItems(10) = "" & TabTemp!CST_COFINS
      item.SubItems(11) = "" & Format(TabTemp!ALIQUOTA_COFINS, strFormatacao2Digitos)
      item.SubItems(12) = "" & Format(TabTemp!perc_base_reduz, strFormatacao2Digitos)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_CFOP"
End Sub

Private Sub LIMPA_CFOP()
'On Error GoTo ERRO_TRATA

   txtFISCO.Text = ""
   txtCFOP_ID.Text = ""
   txtDesc_CFOP.Text = ""
   txtICMS_Dentro.Text = ""
   txtIPI.Text = ""
   txtISS.Text = ""
   txtSUBST.Text = ""
   txtBaseReduz.Text = ""
   txtICMS_Fora.Text = ""
   cmbUFDestino.Text = ""
   cmbCSTICMS.Text = ""
   cmbCSTORIG.Text = ""
   cmbPIS.Text = ""
   txtPIS.Text = ""
   cmbCOFINS.Text = ""
   txtCOFINS.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_CFOP"
End Sub

Private Sub txtCodgGrupo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCodgGrupo.Text <> "" Then
         If IsNumeric(txtCodgGrupo.Text) Then _
            PROCURA_GRUPO_PRODUTO txtCodgGrupo.Text
         Else: txtCodgGrupo.Text = MAX_ID("FAMILIAPRODUTO_ID", "FAMILIAPRODUTO", "", "", "", "")
      End If
      txtDescGrupo.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodgGrupo_KeyPress"
End Sub

Private Sub txtDescGrupo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtUN.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescGrupo_KeyPress"
End Sub

Private Sub txtun_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDescUnidade.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtun_KeyPress"
End Sub

Private Sub txtDescUnidade_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPercVenda.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtun_KeyPress"
End Sub

Private Sub txtPercVenda_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCodgGrupo.SetFocus
      GRAVA_GRUPO_PRODUTO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDescUnidade_KeyPress"
End Sub

Private Sub cmdGrupoLimpa_Click()
'On Error GoTo ERRO_TRATA

   LIMPA_GRUPO_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdGrupoLimpa_Click"
End Sub

Private Sub cmdGrupoMata_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodgGrupo.Text) <> "" Then
      SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " where FAMILIAPRODUTO_ID = " & txtCodgGrupo.Text
      If TabConsulta.State = 1 Then TabConsulta.Close
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         CONECTA_RETAGUARDA.Execute "Delete from FAMILIAPRODUTO Where FAMILIAPRODUTO_ID = " & txtCodgGrupo.Text
         LIMPA_GRUPO_PRODUTO
      End If
      TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdGrupoMata_Click"
End Sub

Private Sub cmdGrupoSair_Click()
'On Error GoTo ERRO_TRATA

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdGrupoSair_Click"
End Sub

Private Sub cmbUFDestino_Click()
   cmbCSTICMS.SetFocus
End Sub

Private Sub cmbUFDestino_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCSTICMS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbUFDestino_KeyPress"
End Sub

Private Sub cmbUFDestino_LostFocus()
   cmbUFDestino.BackColor = &HFFFFFF
End Sub

Private Sub GRAVA_GRUPO_PRODUTO()
'On Error GoTo ERRO_TRATA

   If ((txtCodgGrupo.Text <> "") And (txtDescGrupo.Text <> "")) Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " where codg_familia = '" & Trim(txtCodgGrupo.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         FAMILIA_ID_N = TabTemp.Fields("familiaproduto_id").Value

         SQL = "UPDATE FAMILIAPRODUTO SET "
         SQL = SQL & " Descricao = '" & Trim(txtDescGrupo.Text) & "'"
         SQL = SQL & " , UNIDADE_MEDIDA = '" & Trim(txtUN.Text) & "'"
         SQL = SQL & " , DESC_UNIDADE_MEDIDA = '" & Trim(txtDescUnidade.Text) & "'"
         SQL = SQL & " , producao = " & chkProducao.Value
         SQL = SQL & " , PERC_COMPOE_VENDA = " & tpMOEDA(txtPercVenda.Text)
         SQL = SQL & " where codg_familia = '" & Trim(txtCodgGrupo.Text) & "'"
         Else
            FAMILIA_ID_N = MAX_ID("familiaproduto_id", "familiaPRODUTO", "", "", "", "")

            SQL = "INSERT INTO FAMILIAPRODUTO "
            SQL = SQL & " (FAMILIAPRODUTO_ID,CODG_FAMILIA,DESCRICAO,UNIDADE_MEDIDA,DESC_UNIDADE_MEDIDA,producao,PERC_COMPOE_VENDA) "
            SQL = SQL & " VALUES ("
            SQL = SQL & FAMILIA_ID_N
            SQL = SQL & ",'" & Trim(txtCodgGrupo.Text) & "'"
            SQL = SQL & ",'" & Trim(txtDescGrupo.Text) & "'"
            SQL = SQL & ",'" & Trim(txtUN.Text) & "'"
            SQL = SQL & ",'" & Trim(txtDescUnidade.Text) & "'"
            SQL = SQL & "," & chkProducao.Value
            SQL = SQL & " ," & tpMOEDA(txtPercVenda.Text)
            SQL = SQL & ")"
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

   If TabProduto.State = 1 Then _
      TabProduto.Close

      SQL = "select familiaproduto_id,producao from FAMILIAPRODUTO WITH (NOLOCK)"
      SQL = SQL & " where producao = 1 "
      SQL = SQL & " and codg_familia = '" & Trim(txtCodgGrupo.Text) & "'"
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabProduto.EOF
         NUMR_ID_N = 0
         If Not IsNull(TabProduto.Fields(1).Value) Then
            If TabProduto.Fields(1).Value = True Then
               NUMR_ID_N = 1
               Else: NUMR_ID_N = 0
            End If

            SQL = "update produto set "
            SQL = SQL & " conceder_producao = " & NUMR_ID_N

            SQL = SQL & " where familiaproduto_id = " & TabProduto.Fields(0).Value
            CONECTA_RETAGUARDA.Execute SQL
         End If

         TabProduto.MoveNext
      Wend

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SETA_GRID_GRUPO_PRODUTOS
   LIMPA_GRUPO_PRODUTO
   NUMR_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_GRUPO_PRODUTOk"
End Sub

Private Sub PROCURA_GRUPO_PRODUTO(Codg_Familia_A As String)
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " where codg_familia = '" & Trim(Codg_Familia_A) & "'"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      LIMPA_GRUPO_PRODUTO

      If Not IsNull(TabConsulta.Fields("producao").Value) Then
         If TabConsulta.Fields("producao").Value = True Then
            chkProducao.Value = 1
            Else: chkProducao.Value = 0
         End If
      End If
      txtPercVenda.Text = "" & TabConsulta.Fields("perc_compoe_venda").Value
      FAMILIA_ID_N = 0 & TabConsulta!FAMILIAPRODUTO_ID
      txtCodgGrupo.Text = Trim(TabConsulta.Fields("codg_familia").Value)
      txtDescGrupo.Text = "" & TabConsulta!DESCRICAO
      txtUN.Text = "" & TabConsulta!UNIDADE_MEDIDA
      txtDescUnidade.Text = TabConsulta!DESC_UNIDADE_MEDIDA
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_GRUPO_PRODUTO"
End Sub

Private Sub LIMPA_GRUPO_PRODUTO()
'On Error GoTo ERRO_TRATA

   txtPercVenda.Text = ""
   FAMILIA_ID_N = 0
   txtCodgGrupo.Text = ""
   txtDescGrupo.Text = ""
   txtDescUnidade.Text = ""
   txtUN.Text = ""
   chkProducao.Value = 0

   SETA_GRID_GRUPO_PRODUTOS

   txtCodgGrupo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_GRUPO_PRODUTO"
End Sub

Public Sub SETA_GRID_GRUPO_PRODUTOS()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0

   SQL = "select CODG_FAMILIA,DESCRICAO,UNIDADE_MEDIDA,DESC_UNIDADE_MEDIDA,producao,perc_compoe_venda"
   SQL = SQL & " from FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " order by DESCRICAO "

   adoFamilia.ConnectionString = AUTENTICA_GRID
   adoFamilia.CommandType = adCmdText

   adoFamilia.RecordSource = SQL
   adoFamilia.Enabled = True
   adoFamilia.Refresh

   grdFamilia.Columns(0).DataField = "CODG_FAMILIA"
   grdFamilia.Columns(0).Caption = "Código"
   grdFamilia.Columns(0).Width = 1000

   grdFamilia.Columns(1).DataField = "descricao"
   grdFamilia.Columns(1).Caption = "Descrição"
   grdFamilia.Columns(1).Width = 4000

   grdFamilia.Columns(2).DataField = "UNIDADE_MEDIDA"
   grdFamilia.Columns(2).Caption = "UN"
   grdFamilia.Columns(2).Width = 800

   grdFamilia.Columns(3).DataField = "DESC_UNIDADE_MEDIDA"
   grdFamilia.Columns(3).Caption = "Unidade"
   grdFamilia.Columns(3).Width = 2000

   grdFamilia.Columns(4).DataField = "producao"
   grdFamilia.Columns(4).Caption = "Procução"
   grdFamilia.Columns(4).Width = 1

   grdFamilia.Columns(5).DataField = "perc_compoe_venda"
   grdFamilia.Columns(5).Caption = "PercCompoeVenda"
   grdFamilia.Columns(5).Width = 3000
   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_GRUPO_PRODUTOS"
End Sub

Private Function RetornaDescricaoCFOP(CfopId As String) As String
'On Error GoTo ERRO_TRATA

    RetornaDescricaoCFOP = ""
    
    'On Error GoTo ERRO_TRATA
    
    SQL = "select * from CFOP WITH (NOLOCK)"
    SQL = SQL & " where cfop_id = '" & CfopId & "'"
    SQL = SQL & " order by cfop_id "
    If TabCFOP.State = 1 Then TabCFOP.Close
    TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

    If Not TabCFOP.EOF Then
        TabCFOP.MoveFirst
        RetornaDescricaoCFOP = TabCFOP!DESCRICAO
    End If
    TabCFOP.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RetornaDescricaoCFOP"
End Function

Private Sub preencheComboCfop(NomeCombo As ComboBox)
'On Error GoTo ERRO_TRATA

    If TabCFOP.State = 1 Then _
      TabCFOP.Close

    SQL = "select * from CFOP WITH (NOLOCK) "
    SQL = SQL & " order by cfop_id"
    TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

    NomeCombo.Clear
    If Not TabCFOP.EOF Then
       'Mundando o ponteiro do mouse, para mostrar para o usuario que esta processando...
        Screen.MousePointer = vbHourglass
        
        TabCFOP.MoveFirst
        Do Until TabCFOP.EOF
            'Importantissimo
            DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"
            
            NomeCombo.AddItem TabCFOP!CFOP_ID & "-" & TabCFOP!DESCRICAO
            TabCFOP.MoveNext
        Loop
        
        'Agora vai ficar boca de porco porque? para economizar
        'leitura no banco de dados, pois com o recordset ja preenchido com os
        'dados eu preencho todos os combos
        'TabCFOP.MoveFirst
        'Do Until TabCFOP.EOF
        '    xxxxx.AddItem TabCFOP!cfop_id & "-" & TabCFOP!Descricao
        'Loop
    End If
    
    'Voltando o ponteiro do mouse para o tipo default, ponteiro.
    Screen.MousePointer = vbDefault
    TabCFOP.Close
    
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "preencheComboCfop"
End Sub

Sub PROCESSA_REGISTRO()
'On Error GoTo ERRO_TRATA

   If Trim(cmbCCAux.Text) = "" Then _
      cmbCCAux.Text = 0

   If cmbAuxForma.Text = "" Then
      MsgBox "Selecione forma pagamento."
      cmbForma.SetFocus
      Exit Sub
   End If
   If txtTIPOVENDA.Text = "" Then
      MsgBox "Código inválido."
      txtTIPOVENDA.SetFocus
      Exit Sub
   End If
   If txtDESCTIPOVENDA.Text = "" Then
      MsgBox "Descrição inválida."
      txtDESCTIPOVENDA.SetFocus
      Exit Sub
   End If

   If txtParcela.Text = "" Then _
      txtParcela.Text = 0
   If txtDiasPrazo.Text = "" Then _
      txtDiasPrazo.Text = 0
   If txtDiaVencto.Text = "" Then _
      txtDiaVencto.Text = 0
   If txtPercJuros.Text = "" Then _
      txtPercJuros.Text = 0

   If Trim(txtDebito.Text) = "" Then _
      txtDebito.Text = 0
   If Not IsNumeric(txtDebito.Text) Then _
      txtDebito.Text = 0

   If Trim(txtCredito.Text) = "" Then _
      txtCredito.Text = 0
   If Not IsNumeric(txtCredito.Text) Then _
      txtCredito.Text = 0

Dim ADM_CARTAO_ID As Long

ADM_CARTAO_ID = 0 & cmbADMCartaoAUX.Text

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TIPOVENDA WITH (NOLOCK)"
   SQL = SQL & " where tipovenda_id = " & txtTIPOVENDA.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "INSERT INTO TIPOVENDA "
         SQL = SQL & " (EMPRESA_ID, TIPOvenda_id , Descricao, "
         SQL = SQL & " parcela, prazo, formapagto_id, PERC_JUROS,contabiliza,"
         SQL = SQL & " PERC_CARTAO_DEBITO,PERC_CARTAO_credito,permite_desconto,cc_id,PreFatura,"
         SQL = SQL & " pagar,receber,PermiteParcelar,cartaoadm_id,DIAVENCTO)"
      SQL = SQL & " VALUES ("
         SQL = SQL & EMPRESA_ID_N                                          'EMPRESA_ID
         SQL = SQL & "," & Replace(Trim(txtTIPOVENDA.Text), ",", ".")      'TIPOvenda_id
         SQL = SQL & ",'" & Replace(Trim(txtDESCTIPOVENDA.Text), ",", ".") 'Descricao
         SQL = SQL & "'," & Replace(Trim(txtParcela.Text), ",", ".")       'parcela
         SQL = SQL & "," & Replace(Trim(txtDiasPrazo.Text), ",", ".")      'prazo
         SQL = SQL & "," & Replace(Trim(cmbAuxForma.Text), ",", ".")       'formapagto_id
         SQL = SQL & "," & Replace(txtPercJuros.Text, ",", ".")            'PERC_JUROS
         SQL = SQL & "," & chkContabiliza.Value                            'contabiliza
         SQL = SQL & "," & tpMOEDA(txtDebito.Text)                         'PERC_CARTAO_DEBITO
         SQL = SQL & "," & tpMOEDA(txtCredito.Text)                        'PERC_CARTAO_credito
         SQL = SQL & "," & chkPermite_Desconto.Value                       'permite_desconto
         SQL = SQL & "," & Trim(cmbCCAux.Text)
         SQL = SQL & "," & chkPreFatura.Value                              'PreFatura
         SQL = SQL & "," & chkPagar.Value                                  'pagar
         SQL = SQL & "," & chkReceber.Value                                'receber
         SQL = SQL & "," & chkPermiteParcelar.Value                        'PermiteParcelar
         If ADM_CARTAO_ID <= 0 Then
            SQL = SQL & ",Null"
            Else: SQL = SQL & "," & ADM_CARTAO_ID
         End If
         SQL = SQL & "," & Replace(Trim(txtDiaVencto.Text), ",", ".")      'DIAVENCTO
      SQL = SQL & ")"
      Else
         If TabTemp.State = 1 Then _
            TabTemp.Close
      
         SQL = "UPDATE TIPOVENDA SET "
         SQL = SQL & " Empresa_id = " & EMPRESA_ID_N
         SQL = SQL & ", Descricao = '" & Replace(Trim(txtDESCTIPOVENDA.Text), ",", ".") & "'"
         SQL = SQL & ", parcela = " & Replace(Trim(txtParcela.Text), ",", ".")
         SQL = SQL & ", PRAZO = " & Replace(Trim(txtDiasPrazo.Text), ",", ".")
         SQL = SQL & ", DIAVENCTO = " & Replace(Trim(txtDiaVencto.Text), ",", ".")
         SQL = SQL & ", formapagto_id= " & Replace(Trim(cmbAuxForma.Text), ",", ".")
         SQL = SQL & ", PERC_JUROS = " & Replace(Trim(txtPercJuros.Text), ",", ".")
         SQL = SQL & ", contabiliza = " & chkContabiliza.Value
         SQL = SQL & ", permite_desconto = " & chkPermite_Desconto.Value
         SQL = SQL & ", PreFatura = " & chkPreFatura.Value
         SQL = SQL & ", PERC_CARTAO_DEBITO = " & tpMOEDA(txtDebito.Text)
         SQL = SQL & ", PERC_CARTAO_credito = " & tpMOEDA(txtCredito.Text)
         SQL = SQL & ", cc_id = " & Trim(cmbCCAux.Text)
         SQL = SQL & ", pagar = " & chkPagar.Value
         SQL = SQL & ", receber = " & chkReceber.Value
         SQL = SQL & ", PermiteParcelar = " & chkPermiteParcelar.Value

         If ADM_CARTAO_ID <= 0 Then
            SQL = SQL & ", cartaoadm_id = Null"
            Else: SQL = SQL & ", cartaoadm_id = " & ADM_CARTAO_ID
         End If

         SQL = SQL & " where tipovenda_id = " & Replace(Trim(txtTIPOVENDA.Text), ",", ".")
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID_VENDA
   LIMPA_TIPOVENDA
   txtTIPOVENDA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_REGISTRO"
End Sub

Private Sub SETA_GRID_PAGTO()
'On Error GoTo ERRO_TRATA

   LISTAPAGTO.ListItems.Clear
   'MODALIDADE
   cmbAuxForma.Clear
   cmbForma.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " where FORMAPAGTO_ID < 9999 "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = LISTAPAGTO.ListItems.Add(, "seq." & TabTemp!FORMAPAGTO_ID, TabTemp!FORMAPAGTO_ID)
      item.SubItems(1) = "" & TabTemp!DESCRICAO
      item.SubItems(2) = "" & TabTemp.Fields("status").Value

      If Not IsNull(TabTemp.Fields("contab_balcao").Value) Then
         If TabTemp.Fields("contab_balcao").Value = True Then
            item.SubItems(3) = "Sim"
            Else: item.SubItems(3) = "Não"
         End If
      End If

      If Not IsNull(TabTemp.Fields("baixaauto").Value) Then
         If TabTemp.Fields("baixaauto").Value = True Then
            item.SubItems(4) = "Sim"
            Else: item.SubItems(4) = "Não"
         End If
      End If

      If Not IsNull(TabTemp.Fields("contab_tesora").Value) Then
         If TabTemp.Fields("contab_tesora").Value = True Then
            item.SubItems(5) = "Sim"
            Else: item.SubItems(5) = "Não"
         End If
      End If

      cmbForma.AddItem TabTemp!DESCRICAO
      cmbAuxForma.AddItem TabTemp!FORMAPAGTO_ID

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_PAGTO"
End Sub
'============================== CFOP
'============================== CFOP
'============================== CFOP
'============================== CFOP
'============================== CFOP
'============================== CFOP
'============================== CFOP
'============================== CFOP
Private Sub cmdCFOPLIMPA_Click()
'On Error GoTo ERRO_TRATA

   LIMPA_CFOP
   SETA_GRID_CFOP
   txtCFOP_ID.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCFOPLIMPA_Click"
End Sub

Private Sub cmdCFOPMATA_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtCFOP_ID.Text) <> "" And Trim(txtUFORIGEM.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CFOP WITH (NOLOCK)"
      SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         NUMR_ID_N = 0

         If TabNOTA.State = 1 Then _
            TabNOTA.Close
         SQL = "select cfopuf_id from CFOPUF WITH (NOLOCK)"
         SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
         SQL = SQL & " and uf_origem = '" & Trim(txtUFORIGEM.Text) & "'"
         If Trim(cmbUFDestino.Text) <> "" Then
         SQL = SQL & " and uf_destino = '" & Trim(cmbUFDestino.Text) & "'"
         End If
         TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabNOTA.EOF Then _
            NUMR_ID_N = 0 & TabNOTA.Fields(0).Value
         If TabNOTA.State = 1 Then _
            TabNOTA.Close

         If NUMR_ID_N > 0 Then
            Msg = "Confirma exclusão ?"
            PERGUNTA Msg, vbYesNo + 32, "Cfop Mata", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               SQL = "delete ALIQUOTA_UF where cfopuf_id = " & NUMR_ID_N
               CONECTA_RETAGUARDA.Execute SQL

               SQL = "delete CFOPUF where cfopuf_id = " & NUMR_ID_N
               CONECTA_RETAGUARDA.Execute SQL

               LIMPA_CFOP
               SETA_GRID_CFOP
               txtCFOP_ID.SetFocus
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCFOPMATA_Click"
End Sub

Private Sub cmdCFOPSAIR_Click()
'On Error GoTo ERRO_TRATA

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCFOPSAIR_Click"
End Sub

Private Sub TXTCFOP_ID_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtCFOP_ID.Text) <> "" Then _
         MOSTRA_CFOP_CFOPUF_ALIQUOTA_UF
      txtDesc_CFOP.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCFOP_ID_KeyPress"
End Sub

Private Sub cmbcfopsd_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbcfopsf.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfopsd_KeyPress"
End Sub

Private Sub cmbcfopsf_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbcfoped.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfopsf_KeyPress"
End Sub

Private Sub cmbcfoped_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbcfopef.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfoped_KeyPress"
End Sub

Private Sub cmbcfopef_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     KeyAscii = 0
     cmbdvsd.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbcfopef_KeyPress"
End Sub



Private Sub txtfenapa_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

  If KeyAscii = 13 Then
     GRAVA_PERC_CONT
     cmbcfopsd.SetFocus
  End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtfenapa_KeyPress"
End Sub

Private Sub TXTDESC_CFOP_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Trim(txtCFOP_ID.Text) <> "" And Trim(txtDesc_CFOP.Text) <> "" Then
         KeyAscii = 0
         cmbUFDestino.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTDESC_CFOP_KeyPress"
End Sub

Private Sub txtfisco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_CFOP
      txtCFOP_ID.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtfisco_KeyPress"
End Sub

Private Sub txtICMS_Dentro_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtICMS_Fora.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_Dentro_KeyPress"
End Sub
Private Sub txtipi_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtSUBST.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtipi_KeyPress"
End Sub
Private Sub txtiss_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtFISCO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtiss_KeyPress"
End Sub

Private Sub txtsubst_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtFISCO.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtsubst_KeyPress"
End Sub

Private Sub txtICMS_Fora_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCSTORIG.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtICMS_Fora_KeyPress"
End Sub

Private Sub cmbPIS_Click()
   txtPIS.SetFocus
End Sub

Private Sub cmbPIS_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtPIS.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbPIS_KeyPress"
End Sub

'===========================
Private Sub GRAVA_CFOP()
'On Error GoTo ERRO_TRATA

   If Trim(txtCFOP_ID.Text) <> "" And Trim(txtDesc_CFOP.Text) <> "" Then
      If TabCFOP.State = 1 Then _
         TabCFOP.Close

      SQL = "select * from CFOP WITH (NOLOCK)"
      SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
      TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCFOP.EOF Then
         SQL = "INSERT INTO CFOP "
            SQL = SQL & "(cfop_id,estabelecimento_id,Descricao,perc_ipi,Perc_iss,"
            SQL = SQL & "Perc_subst_trib) "
         SQL = SQL & " VALUES ("
            SQL = SQL & Trim(txtCFOP_ID.Text)                     'CFOP_ID  nvarchar(10)
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                'ESTABELECIMENTO_ID   int
            SQL = SQL & ",'" & Trim(txtDesc_CFOP.Text) & "'"      'DESCRICAO   nvarchar(100)
            SQL = SQL & "," & Trim(txtIPI.Text)                   'PERC_IPI int
            SQL = SQL & "," & Trim(txtISS.Text)                   'PERC_ISS int
            SQL = SQL & "," & Trim(txtSUBST.Text)                 'PERC_SUBST_TRIB   int
            SQL = SQL & ",'" & Trim(txtFISCO.Text) & "'"          'MSGFISCO nvarchar(100)  Checked
         SQL = SQL & ")"
         Else
            SQL = "UPDATE CFOP SET "
               SQL = SQL & " Descricao = '" & Trim(txtDesc_CFOP.Text) & "'"
               SQL = SQL & ", perc_ipi = " & Trim(txtIPI.Text)
               SQL = SQL & ", Perc_iss = " & Trim(txtISS.Text)
               SQL = SQL & ", Perc_subst_trib = " & Trim(txtSUBST.Text)
               SQL = SQL & ", MSGFISCO = '" & Trim(txtFISCO.Text) & "'"
            SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
            'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      End If
      If TabCFOP.State = 1 Then _
         TabCFOP.Close

      CONECTA_RETAGUARDA.Execute SQL

'========CFOPUF
      Dim CFOPUF_ID_N As Integer

      SQL = "select * from CFOPUF WITH (NOLOCK)"
      SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
      SQL = SQL & " and uf_origem = '" & Trim(txtUFORIGEM.Text) & "'"
      SQL = SQL & " and uf_destino = '" & Trim(cmbUFDestino.Text) & "'"
      TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCFOP.EOF Then
         CFOPUF_ID_N = 0 & MAX_ID("CFOPUF_id", "CFOPUF", "", "", "", "")

         SQL = "INSERT INTO CFOPuf "
            SQL = SQL & "(cfopuf_id,cfop_id,uf_origem,uf_destino) "
         SQL = SQL & " VALUES ("
            SQL = SQL & CFOPUF_ID_N
            SQL = SQL & ",'" & Trim(txtCFOP_ID.Text) & "'"
            SQL = SQL & ",'" & Trim(txtUFORIGEM.Text) & "'"
            SQL = SQL & ",'" & Left(cmbUFDestino.Text, 2) & "'"
         SQL = SQL & ")"
         Else
            CFOPUF_ID_N = 0 & TabCFOP.Fields("CFOPUF_ID").Value

            SQL = "UPDATE CFOPuf SET "
               SQL = SQL & " CFOP_ID = '" & Trim(txtCFOP_ID.Text) & "'"
               SQL = SQL & ",uf_origem = '" & Trim(txtUFORIGEM.Text) & "'"
               SQL = SQL & ",uf_destino = '" & Left(cmbUFDestino.Text, 2) & "'"
            SQL = SQL & " where cfopuf_id = " & CFOPUF_ID_N
      End If
      If TabCFOP.State = 1 Then _
         TabCFOP.Close

      CONECTA_RETAGUARDA.Execute SQL

'============ALIQUOTA_UF
      SQL = "select * from ALIQUOTA_UF WITH (NOLOCK)"
      SQL = SQL & " where cfopuf_id = " & CFOPUF_ID_N
      TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCFOP.EOF Then
         SQL = "insert into ALIQUOTA_UF "
            SQL = SQL & "(cfopuf_id,CST_ICMS,ALIQUOTA_ICMS_DENTRO,"
            SQL = SQL & " ALIQUOTA_ICMS_FORA,CST_PIS,ALIQUOTA_PIS,CST_COFINS,ALIQUOTA_COFINS,"
            SQL = SQL & " PERC_BASE_REDUZ,CST_ORIG_ICMS) "
         SQL = SQL & " VALUES ("
            SQL = SQL & CFOPUF_ID_N
            SQL = SQL & ",'" & Left(cmbCSTICMS.Text, 3) & "'"        'CST_ICMS

            SQL = SQL & ",'" & tpMOEDA(txtICMS_Dentro.Text) & "'"    'ALIQUOTA_ICMS_DENTRO
            SQL = SQL & ",'" & tpMOEDA(txtICMS_Fora.Text) & "'"      'ALIQUOTA_ICMS_FORA

            SQL = SQL & ",'" & Left(cmbPIS.Text, 2) & "'"            'CST_PIS

            SQL = SQL & ",'" & tpMOEDA(txtPIS.Text) & "'"            'ALIQUOTA_PIS

            SQL = SQL & ",'" & Left(cmbCOFINS.Text, 2) & "'"         'CST_COFINS

            SQL = SQL & ",'" & tpMOEDA(txtCOFINS.Text) & "'"         'ALIQUOTA_COFINS

            SQL = SQL & ",'" & tpMOEDA(txtBaseReduz.Text) & "'"   'PERC_BASE_REDUZ
            SQL = SQL & ",'" & Left(cmbCSTORIG.Text, 1) & "'"       'CST_ORIG_ICMS
         SQL = SQL & ")"
         Else
            SQL = "UPDATE ALIQUOTA_UF SET "
               SQL = SQL & " CST_ICMS = '" & Left(cmbCSTICMS.Text, 3) & "'"                  'CST_ICMS

               SQL = SQL & ",ALIQUOTA_ICMS_DENTRO = '" & tpMOEDA(txtICMS_Dentro.Text) & "'"  'ALIQUOTA_ICMS_DENTRO
               SQL = SQL & ",ALIQUOTA_ICMS_FORA = '" & tpMOEDA(txtICMS_Fora.Text) & "'"      'ALIQUOTA_ICMS_FORA

               SQL = SQL & ",CST_PIS = '" & Left(cmbPIS.Text, 2) & "'"                       'CST_PIS

               SQL = SQL & ",ALIQUOTA_PIS = '" & tpMOEDA(txtPIS.Text) & "'"                  'ALIQUOTA_PIS

               SQL = SQL & ",CST_COFINS = '" & Left(cmbCOFINS.Text, 2) & "'"                 'CST_COFINS

               SQL = SQL & ",ALIQUOTA_COFINS = '" & tpMOEDA(txtCOFINS.Text) & "'"            'ALIQUOTA_COFINS

               SQL = SQL & ", PERC_BASE_REDUZ = '" & tpMOEDA(txtBaseReduz.Text) & "'"
               SQL = SQL & ", CST_ORIG_ICMS = '" & Left(cmbCSTORIG.Text, 1) & "'"
            SQL = SQL & " where cfopuf_id = " & CFOPUF_ID_N
      End If
      If TabCFOP.State = 1 Then _
         TabCFOP.Close

      CONECTA_RETAGUARDA.Execute SQL

'----------------
      LIMPA_CFOP
      SETA_GRID_CFOP
      txtCFOP_ID.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CFOP"
End Sub

Private Sub cmdGravarBalanca_Click()
'On Error GoTo ERRO_TRATA

   If Trim(txtCasaInicioCodgProdBarra.Text) = "" Then _
      txtCasaInicioCodgProdBarra.Text = 0
   If Trim(txtTamanhoCodgProdBarra.Text) = "" Then _
      txtTamanhoCodgProdBarra.Text = 0
   If Trim(txtTamanhoPesoValorBarra.Text) = "" Then _
      txtTamanhoPesoValorBarra.Text = 0
   If Trim(txtCodgProdutoReserva.Text) = "" Then _
      txtCodgProdutoReserva.Text = 1

   SQL = "update ESTABELECIMENTO set "
   SQL = SQL & " TamanhoCodgProdBarra = " & Trim(txtTamanhoCodgProdBarra.Text)
   SQL = SQL & ",TamanhoPesoValorBarra = " & Trim(txtTamanhoPesoValorBarra.Text)
   SQL = SQL & ",CasaInicioCodgProdBarra = " & Trim(txtCasaInicioCodgProdBarra.Text)
   If chkPanific.Value = 0 Then
      SQL = SQL & ",INDR_PANIFIC = 'false'"
      Else: SQL = SQL & ",INDR_PANIFIC = 'true'"
   End If
   If optGramas.Value = True Then
      SQL = SQL & ",peso_valor = '" & optGramas.Caption & "'"
      Else: SQL = SQL & ",peso_valor = '" & optValor.Caption & "'"
   End If
   SQL = SQL & ",CODG_PROD_RESERVA = " & txtCodgProdutoReserva.Text
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   MsgBox "Processo realizado com sucesso."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdGravarBalanca_Click"
End Sub

Sub MOSTRA_BALANCA()
'On Error GoTo ERRO_TRATA

   optValor.Value = False
   optGramas.Value = False

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select TamanhoCodgProdBarra,TamanhoPesoValorBarra,INDR_PANIFIC,peso_valor,CODG_PROD_RESERVA,CasaInicioCodgProdBarra "
   SQL = SQL & " from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If Not IsNull(TabTemp.Fields("CasaInicioCodgProdBarra").Value) Then _
         txtCasaInicioCodgProdBarra.Text = TabTemp.Fields("CasaInicioCodgProdBarra").Value

      If Not IsNull(TabTemp.Fields("TamanhoCodgProdBarra").Value) Then _
         txtTamanhoCodgProdBarra.Text = TabTemp.Fields("TamanhoCodgProdBarra").Value

      If Not IsNull(TabTemp.Fields("TamanhoPesoValorBarra").Value) Then _
         txtTamanhoPesoValorBarra.Text = TabTemp.Fields("TamanhoPesoValorBarra").Value

      If Not IsNull(TabTemp.Fields("INDR_PANIFIC").Value) Then
         If TabTemp.Fields("INDR_PANIFIC").Value = 0 Then
            chkPanific.Value = 0
            Else: chkPanific.Value = 1
         End If
      End If
      If Not IsNull(TabTemp.Fields("peso_valor").Value) Then _
         If Trim(UCase(TabTemp.Fields("peso_valor").Value)) = UCase("valor") Then _
            optValor.Value = True
         If Trim(UCase(TabTemp.Fields("peso_valor").Value)) = UCase("gramas") Then _
            optGramas.Value = True

      If Not IsNull(TabTemp.Fields("CODG_PROD_RESERVA").Value) Then _
         txtCodgProdutoReserva.Text = TabTemp.Fields("CODG_PROD_RESERVA").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_BALANCA"
End Sub

Private Sub SETA_GRID_CSOSN()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0

   SQL = "select * from CSOSN WITH (NOLOCK)"

   adoCSOSN.ConnectionString = AUTENTICA_GRID
   adoCSOSN.CommandType = adCmdText

   adoCSOSN.RecordSource = SQL
   adoCSOSN.Enabled = True
   adoCSOSN.Refresh

   Set txtCodgCSOSN.DataSource = adoCSOSN
   txtCodgCSOSN.DataField = "codigo"

   Set txtDescCSOSN.DataSource = adoCSOSN
   txtDescCSOSN.DataField = "descricao"

   Set txtInstrucaoCSOSN.DataSource = adoCSOSN
   txtInstrucaoCSOSN.DataField = "obs"

   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_CSOSN"
End Sub

Private Sub SETA_GRID_CST()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0

   SQL = "select * from CST WITH (NOLOCK)"

   adoCST.ConnectionString = AUTENTICA_GRID
   adoCST.CommandType = adCmdText

   adoCST.RecordSource = SQL
   adoCST.Enabled = True
   adoCST.Refresh

   Set txtCodgCST.DataSource = adoCST
   txtCodgCST.DataField = "codigo"

   Set txtDescCST.DataSource = adoCST
   txtDescCST.DataField = "descricao"

   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_CST"
End Sub

Sub MOSTRA_CFOP_CFOPUF_ALIQUOTA_UF()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CFOP WITH (NOLOCK)"
   SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtDesc_CFOP.Text = "" & Trim(TabTemp!DESCRICAO)
      txtIPI.Text = "" & TabTemp!PERC_IPI
      txtISS.Text = "" & TabTemp!PERC_ISS
      txtSUBST.Text = "" & TabTemp!PERC_SUBST_TRIB
      txtFISCO.Text = "" & TabTemp!MSGFISCO
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CFOPUF WITH (NOLOCK)"
   SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
   SQL = SQL & " and uf_origem = '" & Trim(txtUFORIGEM.Text) & "'"

   If Trim(cmbUFDestino.Text) <> "" Then _
      SQL = SQL & " and uf_destino = '" & Trim(cmbUFDestino.Text) & "'"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      NUMR_ID_N = 0 & TabTemp.Fields(0).Value
      cmbUFDestino.Text = "" & TabTemp.Fields("UF_DESTINO").Value
      txtUFORIGEM.Text = "" & TabTemp.Fields("UF_ORIGEM").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from ALIQUOTA_UF WITH (NOLOCK)"
   SQL = SQL & " where cfopUF_id = " & NUMR_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      cmbCSTICMS.Text = "" & TabTemp.Fields("CST_ICMS").Value
      txtICMS_Dentro.Text = "" & TabTemp.Fields("ALIQUOTA_ICMS_DENTRO").Value
      txtICMS_Fora.Text = "" & TabTemp.Fields("ALIQUOTA_ICMS_FORA").Value
      cmbPIS.Text = "" & TabTemp.Fields("CST_PIS").Value
      txtPIS.Text = "" & TabTemp.Fields("ALIQUOTA_PIS").Value
      cmbCOFINS.Text = "" & TabTemp.Fields("CST_COFINS").Value
      txtCOFINS.Text = "" & TabTemp.Fields("ALIQUOTA_COFINS").Value
      txtBaseReduz.Text = "" & TabTemp.Fields("PERC_BASE_REDUZ").Value
      cmbCSTORIG.Text = "" & TabTemp.Fields("cst_orig_icms").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   NUMR_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CFOP_CFOPUF_ALIQUOTA_UF"
End Sub

Sub MOSTRA_CFOP_CFOPUF_ALIQUOTA_UF_DO_GRID(CFOPUF_ID_N As Long)
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0
   LIMPA_CFOP

   If TabTemp.State = 1 Then _
      TabTemp.Close

'MsgBox Trim(lstCFOP.SelectedItem.ListSubItems.item(6).Text)

   SQL = "select * from CFOPUF WITH (NOLOCK)"
   SQL = SQL & " where cfopuf_id = " & CFOPUF_ID_N

   'If Trim(lstCFOP.SelectedItem.ListSubItems.item(3).Text) <> "" Then _
      SQL = SQL & " and uf_origem = '" & Trim(lstCFOP.SelectedItem.ListSubItems.item(3).Text) & "'"
   'If Trim(lstCFOP.SelectedItem.ListSubItems.item(6).Text) <> "" Then _
      SQL = SQL & " and uf_destino = '" & Trim(lstCFOP.SelectedItem.ListSubItems.item(6).Text) & "'"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtCFOP_ID.Text = "" & TabTemp.Fields("cfop_id").Value
      txtUFORIGEM.Text = "" & TabTemp.Fields("uf_origem").Value
      cmbUFDestino.Text = "" & TabTemp.Fields("uf_destino").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CFOP WITH (NOLOCK)"
   SQL = SQL & " where cfop_id = '" & Trim(txtCFOP_ID.Text) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtDesc_CFOP.Text = "" & Trim(TabTemp!DESCRICAO)
      txtIPI.Text = "" & TabTemp!PERC_IPI
      txtISS.Text = "" & TabTemp!PERC_ISS
      txtSUBST.Text = "" & TabTemp!PERC_SUBST_TRIB
      txtFISCO.Text = "" & TabTemp!MSGFISCO
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from ALIQUOTA_UF WITH (NOLOCK)"
   SQL = SQL & " where cfopUF_id = " & CFOPUF_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      cmbCSTICMS.Text = "" & TabTemp.Fields("CST_ICMS").Value
      txtICMS_Dentro.Text = "" & TabTemp.Fields("ALIQUOTA_ICMS_DENTRO").Value
      txtICMS_Fora.Text = "" & TabTemp.Fields("ALIQUOTA_ICMS_FORA").Value
      cmbPIS.Text = "" & TabTemp.Fields("CST_PIS").Value
      txtPIS.Text = "" & TabTemp.Fields("ALIQUOTA_PIS").Value
      cmbCOFINS.Text = "" & TabTemp.Fields("CST_COFINS").Value
      txtCOFINS.Text = "" & TabTemp.Fields("ALIQUOTA_COFINS").Value
      txtBaseReduz.Text = "" & TabTemp.Fields("PERC_BASE_REDUZ").Value
      cmbCSTORIG.Text = "" & TabTemp.Fields("cst_orig_icms").Value
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   NUMR_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_CFOP_CFOPUF_ALIQUOTA_UF_DO_GRID"
End Sub
