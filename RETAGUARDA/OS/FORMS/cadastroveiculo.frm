VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCADASTROVEICULO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Veículo"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "CADASTROVEICULO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPlaca 
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbAuxCombustivel 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   28
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbCombustivel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox cmbAuxCor 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbCor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtKm 
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtMotor 
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
      Left            =   840
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtDescricao 
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
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.ComboBox cmbAuxTipo 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbTIPO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtMODELO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtANO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   11
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox txtCHASSI 
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
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1320
      Width           =   5055
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16777215
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   0
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
            Picture         =   "CADASTROVEICULO.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVEICULO.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVEICULO.frx":077C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVEICULO.frx":0BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVEICULO.frx":1024
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVEICULO.frx":1344
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROVEICULO.frx":1798
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   1111
      ButtonWidth     =   1111
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "importa"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSMask.MaskEdBox txtCGCCPF 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView LISTACHASSI 
      Height          =   2865
      Left            =   0
      TabIndex        =   29
      Top             =   3600
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   5054
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
      ForeColor       =   0
      BackColor       =   16777152
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chassi"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Placa"
         Object.Width           =   1960
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "ANO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "MODELO"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "TIPO"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combustível:"
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
      TabIndex        =   27
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor:"
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
      Left            =   5160
      TabIndex        =   25
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km:"
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
      Left            =   360
      TabIndex        =   24
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motor:"
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
      TabIndex        =   23
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição/Modelo:"
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
      Left            =   2160
      TabIndex        =   22
      Top             =   840
      Width           =   1995
   End
   Begin VB.Label Label6 
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
      TabIndex        =   13
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Veículo:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Dt.Cadastro:"
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
      Left            =   7200
      TabIndex        =   19
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
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
      Left            =   1920
      TabIndex        =   17
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label lblCpf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassi:"
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
      Left            =   2640
      TabIndex        =   14
      Top             =   1320
      Width           =   780
   End
End
Attribute VB_Name = "frmCADASTROVEICULO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   ABRE_BANCO_AUXILIAR
   SETA_GRID_CHASSI
   cmbAuxTipo.Clear
   cmbtipo.Clear
   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'P' "
   SQL = SQL & "order by desc_a"
   Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
   While Not TABDESCR.EOF
      cmbtipo.AddItem Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
      cmbAuxTipo.AddItem TABDESCR!Codigo
      TABDESCR.MoveNext
   Wend
   TABDESCR.Close

   cmbAuxCor.Clear
   cmbCor.Clear
   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'Q' "
   SQL = SQL & "order by desc_a"
   Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
   While Not TABDESCR.EOF
      cmbCor.AddItem Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
      cmbAuxCor.AddItem TABDESCR!Codigo
      TABDESCR.MoveNext
   Wend
   TABDESCR.Close
   
   cmbAuxCombustivel.Clear
   cmbCombustivel.Clear
   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'S' "
   SQL = SQL & "order by desc_a"
   Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
   While Not TABDESCR.EOF
      cmbCombustivel.AddItem Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
      cmbAuxCombustivel.AddItem TABDESCR!Codigo
      TABDESCR.MoveNext
   Wend
   TABDESCR.Close
   
   txtDtIni.PromptInclude = False
   txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
End Sub

Private Sub Form_Load()
   Call CentralizaJanela2(frmCADASTROVEICULO)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DBARQAUX.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
      Case "importa"
Dim DBInterBase As New ADODB.Connection
Dim TaBInterBase As New ADODB.Recordset
Dim TaBInterBaseTemp As New ADODB.Recordset

'conexão com o IBProvider
'adoConn.Open "provider=LCPI.IBProvider;data source=C:\teste\Employee.gdb;ctype=win1251;user 'id=sysdba;password=masterkey"
'conexão com o SIBProvider
'adoConn.Open "provider=sibprovider;data source=c:\teste\employee.gdb", "sysdba", "masterkey"
'conexão com o IbOleDb Provider
'adoConn.Open "Provider=IbOleDb;Data Source=c:\teste\employee.gdb", "sysdba", "masterkey

'DBInterBase.Open "provider=LCPI.IBProvider;data source=C:\SIOF\Dados\BKP\dboficina.gdb;ctype=win1251;user 'id=sysdba;password=masterkey"

DBInterBase.Open "Provider=IbOleDb;Data Source=C:\SIOF\Dados\BKP\dboficina.gdb", "sysdba", "masterkey"
'DBInterBase.Open "provider=sibprovider;data source=C:\SIOF\Dados\BKP\dboficina.gdb", "sysdba", "masterkey"

SQL = "select * from AUTOMOVEL "
'SQL = "select * from AUTOMOVEL , CONTATO  "
'SQL = SQL & "where AUTOMOVEL.idcontato = CONTATO.idcontato"
TaBInterBase.Open SQL, DBInterBase
'TaBInterBase.Source = SQL
'TaBInterBase.ActiveConnection = DBInterBase
'TaBInterBase.Open

While Not TaBInterBase.EOF
   VALOR_TOTAL_N = 0
   SQL = "select cnpj_cpf from CONTATO "
   SQL = SQL & "where IDCONTATO = " & TaBInterBase!idcontato
   TaBInterBaseTemp.Open SQL, DBInterBase
   If Not TaBInterBaseTemp.EOF Then
      CRITERIO = Replace(TaBInterBaseTemp!cnpj_cpf, "-", "")
      CRITERIO = Replace(CRITERIO, ".", "")
      'MsgBox CRITERIO
      VALOR_TOTAL_N = CRITERIO
      If VALOR_TOTAL_N <= 0 Then _
         VALOR_TOTAL_N = TaBInterBase!idcontato
      Else: VALOR_TOTAL_N = TaBInterBase!idcontato
   End If
   
   SQL = "select cgccpf from CLIENTE "
   SQL = SQL & "where cgccpf = '" & Trim(CRITERIO) & "'"
   Set TABCLI = DBARQEMP.OpenRecordset(SQL, 4)
   If TABCLI.EOF Then
      TABCLI.Close
      'MsgBox "cliente não cadastrado"
      GoTo PULA
   End If
   TABCLI.Close

   SQL = "select * from CHASSI "
   SQL = SQL & "where placa = '" & TaBInterBase!numero_placa & "'"
   Set TABTEMP = DBARQAUX.OpenRecordset(SQL)
   If Not TABTEMP.EOF Then
      TABTEMP.Edit
      Else: TABTEMP.AddNew
   End If
   TABTEMP!CGCCPF = Trim(CRITERIO)
   TABTEMP!placa = TaBInterBase!numero_placa
   If IsNull(TaBInterBase!numero_chassi) Then
      TABTEMP!nr_chassi = Trim(TaBInterBase!numero_placa)
      Else
         If Trim(TaBInterBase!numero_chassi) = "" Then
            TABTEMP!nr_chassi = Trim(TaBInterBase!numero_placa)
            Else: TABTEMP!nr_chassi = TaBInterBase!numero_chassi
         End If
   End If
   TABTEMP!ano = TaBInterBase!ano_fabricacao
   TABTEMP!Modelo = TaBInterBase!ano_modelo
   TABTEMP!Tipo = 0
   'TABTEMP!dt_cad = TaBInterBaseTemp!dt_cadastro
   TABTEMP!DT_CAD = Date
   TABTEMP!Descricao = TaBInterBase!Descricao
   TABTEMP!motor = TaBInterBase!motor
   If Trim(TaBInterBase!km) <> "" Then
      If IsNumeric(Trim(TaBInterBase!km)) Then
         TABTEMP!KM_CADASTRO = Trim(TaBInterBase!km)
      End If
   End If
   TABTEMP!km_atual = Null

   'cadastra cor
   NUMR_SEQ_N = 0
   If Not IsNull(TaBInterBase!cor) Then
      SQL = "select * from DESCR "
      SQL = SQL & "where tipo_a = 'Q' "
      SQL = SQL & " and desc_a = '" & Trim(TaBInterBase!cor) & "'"
      Set TABAUX = DBARQEMP.OpenRecordset(SQL)
      If TABAUX.EOF Then
         NUMR_SEQ_N = 1
         SQL = "select max(codigo) from DESCR "
         SQL = SQL & "where tipo_a = 'Q' "
         Set TABCONSULTA = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TABCONSULTA.EOF Then _
            If Not IsNull(TABCONSULTA.Fields(0).Value) Then _
               NUMR_SEQ_N = 1 + TABCONSULTA.Fields(0).Value
         TABCONSULTA.Close
         TABAUX.AddNew
            TABAUX!TIPO_A = "Q"
            TABAUX!desc_a = TaBInterBase!cor
            TABAUX!Codigo = NUMR_SEQ_N
         TABAUX.Update
         Else: NUMR_SEQ_N = TABAUX!Codigo
      End If
      TABAUX.Close
      TABTEMP!cor = NUMR_SEQ_N
   End If
   'cadastra combustivel
   If Not IsNull(TaBInterBase!idtipo_combustivel) Then
      If TaBInterBase!idtipo_combustivel = 1 Or TaBInterBase!idtipo_combustivel = 4 Or _
         TaBInterBase!idtipo_combustivel = 5 Then
         TABTEMP!combustivel = 2
         Else: TABTEMP!combustivel = TaBInterBase!idtipo_combustivel
      End If
      'SQL = "select * from DESCR "
      'SQL = SQL & "where tipo_a = 'S' "
      'SQL = SQL & " and desc_a = '" & TaBInterBase!idtipo_combustivel & "'"
      'Set TABAUX = DBARQEMP.OpenRecordset(SQL)
      'If TABAUX.EOF Then
      '   TABAUX.AddNew
      '      TABAUX!TIPO_A = "S"
      '      If TaBInterBase!idtipo_combustivel = 3 Then _
      '         TABAUX!desc_a = "Gasolina"
      '      If TaBInterBase!idtipo_combustivel = 2 Then _
      '         TABAUX!desc_a = "Diesel"
      '      If TaBInterBase!idtipo_combustivel = 1 Or TaBInterBase!idtipo_combustivel = 4 Or _
      '         TaBInterBase!idtipo_combustivel = 5 Then _
      '         TABAUX!desc_a = "Disel"
      '      TABAUX!Codigo = TaBInterBase!idtipo_combustivel
      '   TABAUX.Update
      '   Else: NUMR_SEQ_N = TABAUX!Codigo
      'End If
      'TABAUX.Close
   End If
   TABTEMP.Update
   frmCADASTROVEICULO.Caption = VALOR_TOTAL_N
   frmCADASTROVEICULO.Refresh
   TABTEMP.Close
PULA:
   TaBInterBaseTemp.Close
   TaBInterBase.MoveNext
Wend
TaBInterBase.Close
DBInterBase.Close
MsgBox "ok"
      Case "voltar"
         Unload Me
      Case "matar"
         If txtCHASSI.Text <> "" Then
            
            SQL = "select * from CHASSI "
            SQL = SQL & "where placa = '" & txtPlaca.Text & "'"
            Set TABTEMP = DBARQAUX.OpenRecordset(SQL)
            If Not TABTEMP.EOF Then
               SQL = "select * from CABECAOS "
               SQL = SQL & "where placa = '" & TABTEMP!placa & "'"
               Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
               If Not TABAUX.EOF Then
                  TABTEMP.Close
                  TABAUX.Close
                  MsgBox "Impossível excluir chassi, existe O.S. lançada para o mesmo."
                  Exit Sub
               End If
               TABAUX.Close
               Msg = "Confirma Exclusão do chassi ?"
               PERGUNTA
               If RESPOSTA = vbYes Then
                  TABTEMP.Delete
                  LIMPA_CHASSI
                  txtPlaca.SetFocus
               End If
            End If
            TABTEMP.Close
         End If
      Case "gravar"
         GRAVA_CHASSI
         txtPlaca.SetFocus
      Case "limpar"
         LIMPA_CHASSI
      Case "imprimir"
   End Select
End Sub
'==================cgccpf
Private Sub TXTCGCCPF_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Clientes"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   txtCGCCPF.Mask = "##############"
   If CPF_N <> "" Then
      txtCGCCPF.PromptInclude = False
      txtCGCCPF.Text = CPF_N
      CPF_N = ""
   End If
End Sub

Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYCLIENTE.Show 1
         If CPF_N <> "" Then
            txtCGCCPF.Mask = "##############"
            txtCGCCPF.Text = CPF_N
         End If
         CPF_N = ""
   End Select
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      ENDERECO_A = ""
      txtCGCCPF.PromptInclude = False
      If txtCGCCPF.Text = "" Then
         'MsgBox "Informe CNPJ/CPF corretamente"
         txtCGCCPF.Text = "99999999999"
         Else
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
            txtCGCCPF.PromptInclude = False
            CRITERIO = txtCGCCPF.Text
      End If
      txtCGCCPF.PromptInclude = False
      If txtCGCCPF.Text <> "" Then
         CRITERIO = txtCGCCPF.Text
         If Not IsNull(txtCGCCPF.Text) Then
            If Len(txtCGCCPF.Text) <= 11 Then
               txtCGCCPF.Mask = "###.###.###-##"
               Else: txtCGCCPF.Mask = "##.###.###/####-##"
            End If
         End If
         txtCGCCPF.Text = CRITERIO
      End If
      txtCGCCPF.PromptInclude = False
      SQL = "select * from CLIENTE "
      SQL = SQL & "where CGCCPF = '" & txtCGCCPF.Text & "'"
      Set TABCLI = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
      If TABCLI.EOF Then
         Beep
         MsgBox "CPF não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCGCCPF.SetFocus
         Exit Sub
         Else
            If TABCLI!Nome <> "" Then
               txtNome.Text = TABCLI!Nome
               'If Not IsNull(TABCLI!limite_credito) Then _
                  txtLIMITE.Text = Format(TABCLI!limite_credito, "fixed")
               'SQL = "select sum(i.valor_item-i.valor_desconto) from ITEMLANCAMENTO i, LANCAMENTO l "
               'SQL = SQL & "where i.numr_doc = l.numr_doc "
               'SQL = SQL & " and l.prop = '" & TABCLI!CGCCPF & "'"
               'SQL = SQL & " and i.status = 'A' "
               'Set TABAUX = DBARQEMP.OpenRecordset(SQL, 4)
               'If Not TABAUX.EOF Then
               '   If Not IsNull(TABAUX.Fields(0).Value) Then
               '      txtPAGAR.Text = Format(TABAUX.Fields(0).Value, "fixed")
               '      txtPAGAR.Refresh
               '   End If
               'End If
               'TABAUX.Close
            End If
      End If
      GRAVA_CHASSI
      txtPlaca.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub
'======================
Private Sub txtCHASSI_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_CHASSI
      txtKm.SetFocus
   End If
End Sub

Private Sub txtkm_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtANO.SetFocus
   End If
End Sub

Private Sub txtplaca_Change()
   SETA_GRID_CHASSI
End Sub

Private Sub txtPLACA_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmCONSULTACHASSI.Show 1
         If SQL3 <> "" Then
            
            SQL = "select placa from CHASSI "
            SQL = SQL & "where nr_chassi = '" & SQL3 & "'"
            Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
            If Not TABAUX.EOF Then
               txtPlaca.Text = Left(TABAUX!placa, 3) & "-" & Right(TABAUX!placa, 5)
            End If
            TABAUX.Clone
            
         End If
         SQL3 = ""
         txtPlaca.SetFocus
   End Select
End Sub

Private Sub txtplaca_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      MOSTRA_PLACA
      txtDescricao.SetFocus
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

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtMotor.SetFocus
   End If
End Sub

Private Sub txtmotor_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCHASSI.SetFocus
   End If
End Sub

Private Sub txtCHASSI_LostFocus()
   If txtCHASSI.Text = "" Then
      txtCHASSI.Text = txtPlaca.Text
   '   MsgBox "Chassi inválido."
   '   txtCHASSI.SetFocus
   '   Exit Sub
   End If
End Sub

Private Sub txtANO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtMODELO.SetFocus
   End If
End Sub

Private Sub txtMODELO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCor.SetFocus
   End If
End Sub

Private Sub cmbcor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbCombustivel.SetFocus
   End If
End Sub

Private Sub cmbcombustivel_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbtipo.SetFocus
   End If
End Sub

Private Sub cmbTIPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCGCCPF.SetFocus
   End If
End Sub

Private Sub cmbTipo_Click()
   cmbAuxTipo.ListIndex = cmbtipo.ListIndex
End Sub

Private Sub cmbcor_Click()
   cmbAuxCor.ListIndex = cmbCor.ListIndex
End Sub

Private Sub cmbcombustivel_Click()
   cmbAuxCombustivel.ListIndex = cmbCombustivel.ListIndex
End Sub
'=======================
Private Sub MOSTRA_CHASSI()
   If txtCHASSI.Text <> "" Then
      SQL = "select * from CHASSI "
      SQL = SQL & "where nr_chassi = '" & txtCHASSI.Text & "'"
      Set TABTEMP = DBARQAUX.OpenRecordset(SQL)
      If Not TABTEMP.EOF Then
         txtPlaca.Text = Left(TABTEMP!placa, 3) & "-" & Right(TABTEMP!placa, 5)
         txtCHASSI.Text = TABTEMP!nr_chassi
         txtCGCCPF.PromptInclude = False
         txtCGCCPF.Text = TABTEMP!CGCCPF

         SQL = "select * from CLIENTE "
         SQL = SQL & "where CGCCPF = '" & txtCGCCPF.Text & "'"
         Set TABCLI = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
         If Not TABCLI.EOF Then
            txtNome.Text = TABCLI!Nome
         End If
         TABCLI.Close

         If Not IsNull(TABTEMP!ano) Then _
            txtANO.Text = TABTEMP!ano
         If Not IsNull(TABTEMP!Modelo) Then _
            txtMODELO.Text = TABTEMP!Modelo
         If Not IsNull(TABTEMP!Tipo) Then
            SQL = "select * from DESCR "
            SQL = SQL & "where tipo_a = 'P' "
            SQL = SQL & " and codigo = " & TABTEMP!Tipo
            Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TABDESCR.EOF Then _
               cmbtipo.Text = Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
            TABDESCR.Close
            cmbAuxTipo.Text = TABTEMP!Tipo
         End If
      End If
   End If
End Sub

Private Sub MOSTRA_PLACA()
   If txtPlaca.Text <> "" Then
      SQL = "select * from CHASSI "
      SQL = SQL & "where placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
      Set TABTEMP = DBARQAUX.OpenRecordset(SQL)
      If Not TABTEMP.EOF Then
         'txtPlaca.Text = Left(TABTEMP!placa, 3) & "-" & Right(TABTEMP!placa, 5)
         txtCHASSI.Text = TABTEMP!nr_chassi
         'MsgBox Left(TABTEMP!placa, 3) & "-" & Right(TABTEMP!placa, 5)
         txtCGCCPF.PromptInclude = False
         txtCGCCPF.Text = TABTEMP!CGCCPF
            If Not IsNull(TABTEMP!Descricao) Then _
               txtDescricao.Text = TABTEMP!Descricao
            If Not IsNull(TABTEMP!motor) Then _
               txtMotor.Text = TABTEMP!motor
            If Not IsNull(TABTEMP!KM_CADASTRO) Then _
               txtKm.Text = TABTEMP!KM_CADASTRO
            If Not IsNull(TABTEMP!cor) Then
               cmbAuxCor.Text = TABTEMP!cor
               SQL = "select * from DESCR "
               SQL = SQL & "where tipo_a = 'Q' "
               SQL = SQL & " and codigo = " & TABTEMP!cor
               Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
               If Not TABDESCR.EOF Then _
                  cmbCor.Text = Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
               TABDESCR.Close
            End If
            If Not IsNull(TABTEMP!combustivel) Then
               cmbAuxCombustivel.Text = TABTEMP!combustivel
               SQL = "select * from DESCR "
               SQL = SQL & "where tipo_a = 'S' "
               SQL = SQL & " and codigo = " & TABTEMP!combustivel
               Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
               If Not TABDESCR.EOF Then _
                  cmbCombustivel.Text = Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
               TABDESCR.Close
            End If
         SQL = "select * from CLIENTE "
         SQL = SQL & "where CGCCPF = '" & txtCGCCPF.Text & "'"
         Set TABCLI = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
         If Not TABCLI.EOF Then
            txtNome.Text = TABCLI!Nome
         End If
         TABCLI.Close
         If Not IsNull(TABTEMP!ano) Then _
            txtANO.Text = TABTEMP!ano
         If Not IsNull(TABTEMP!Modelo) Then _
            txtMODELO.Text = TABTEMP!Modelo
         If Not IsNull(TABTEMP!Tipo) Then
            SQL = "select * from DESCR "
            SQL = SQL & "where tipo_a = 'P' "
            SQL = SQL & " and codigo = " & TABTEMP!Tipo
            Set TABDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TABDESCR.EOF Then _
               cmbtipo.Text = Trim(TABDESCR!desc_a) & " - " & TABDESCR!Codigo
            TABDESCR.Close
            cmbAuxTipo.Text = TABTEMP!Tipo
         End If
      End If
      
   End If
End Sub

Private Sub LIMPA_CHASSI()
   txtPlaca.Text = ""
   txtDescricao.Text = ""
   txtMotor.Text = ""
   txtCHASSI.Text = ""
   txtKm.Text = ""
   cmbCor.Text = ""
   cmbAuxCor.Text = ""
   cmbAuxCombustivel.Text = ""
   cmbCombustivel.Text = ""
   txtANO.Text = ""
   txtMODELO.Text = ""
   cmbtipo.Text = ""
   cmbAuxTipo.Text = ""
   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = ""
   txtNome.Text = ""
   SETA_GRID_CHASSI
   txtPlaca.SetFocus
End Sub

Private Sub GRAVA_CHASSI()
   If txtPlaca.Text = "" Then
      MsgBox "Número de Placa deve ser informado."
      txtPlaca.SetFocus
      Exit Sub
   End If
   If txtCHASSI.Text = "" Then
      MsgBox "Número de Chassi deve ser informado."
      txtCHASSI.SetFocus
      Exit Sub
   End If
   txtCGCCPF.PromptInclude = False
   If txtCGCCPF.Text = "" Then
      MsgBox "Cliente deve ser informado."
      txtCGCCPF.SetFocus
      Exit Sub
   End If
   If txtCHASSI.Text = "" Then
      MsgBox "Número de Chassi deve ser informado."
      txtCHASSI.SetFocus
      Exit Sub
   End If
   
   SQL = "select * from CHASSI "
   SQL = SQL & "where placa = '" & Replace(Trim(txtPlaca.Text), "-", "") & "'"
   Set TABTEMP = DBARQAUX.OpenRecordset(SQL)
   If Not TABTEMP.EOF Then
      TABTEMP.Edit
      Else: TABTEMP.AddNew
   End If
   TABTEMP!placa = Replace(Trim(txtPlaca.Text), "-", "")
   If txtDescricao.Text <> "" Then
      TABTEMP!Descricao = txtDescricao.Text
      Else: TABTEMP!Descricao = Null
   End If
   If txtMotor.Text <> "" Then
      TABTEMP!motor = Left(txtMotor.Text, 30)
      Else: TABTEMP!motor = Null
   End If
   TABTEMP!nr_chassi = txtCHASSI.Text
   If txtKm.Text <> "" Then
      TABTEMP!KM_CADASTRO = txtKm.Text
      Else: TABTEMP!KM_CADASTRO = Null
   End If
   If cmbAuxCor.Text <> "" Then
      TABTEMP!cor = cmbAuxCor.Text
      Else: TABTEMP!cor = 0
   End If
   If cmbAuxCombustivel.Text <> "" Then
      TABTEMP!combustivel = cmbAuxCombustivel.Text
      Else: TABTEMP!combustivel = 0
   End If
   TABTEMP!CGCCPF = txtCGCCPF.Text
   If txtANO.Text = "" Then
      TABTEMP!ano = 0
      Else: TABTEMP!ano = txtANO.Text
   End If
   If txtMODELO.Text = "" Then
      TABTEMP!Modelo = 0
      Else: TABTEMP!Modelo = txtMODELO.Text
   End If
   If cmbtipo.Text = "" Then
      TABTEMP!Tipo = 0
      Else: TABTEMP!Tipo = cmbAuxTipo.Text
   End If
   If Not IsDate(TABTEMP!DT_CAD) Then _
      TABTEMP!DT_CAD = Date
   TABTEMP.Update
   TABTEMP.Close
   LIMPA_CHASSI
End Sub

Private Sub SETA_GRID_CHASSI()
   NUMR_SEQ_N = 1
   LISTACHASSI.ListItems.Clear
   SQL = "select * from CHASSI "
   SQL = SQL & "where nr_chassi <> '' "
   If txtPlaca.Text <> "" Then _
      SQL = SQL & " and placa like " & Chr$(39) & Replace(txtPlaca.Text, "-", "") & "*" & Chr(39)
   SQL = SQL & " order by ano asc "
   Set TABAUX = DBARQAUX.OpenRecordset(SQL, 4)
   While Not TABAUX.EOF
      Set ITEM = LISTACHASSI.ListItems.Add(, "seq." & TABAUX!placa, TABAUX!nr_chassi)
      ITEM.SubItems(1) = TABAUX!placa
      SQL = "select nome from CLIENTE "
      SQL = SQL & "where cgccpf = '" & TABAUX!CGCCPF & "'"
      Set TABCLI = DBARQEMP.OpenRecordset(SQL, 4)
      If Not TABCLI.EOF Then _
         ITEM.SubItems(2) = TABCLI!Nome
      TABCLI.Close
      If Not IsNull(TABAUX!ano) Then _
         ITEM.SubItems(3) = TABAUX!ano
      If Not IsNull(TABAUX!Modelo) Then _
         ITEM.SubItems(4) = TABAUX!Modelo
      If Not IsNull(TABAUX!Tipo) Then _
         ITEM.SubItems(5) = TABAUX!Tipo
      TABAUX.MoveNext
   Wend
   TABAUX.Close
   txtCGCCPF.PromptInclude = True
End Sub
