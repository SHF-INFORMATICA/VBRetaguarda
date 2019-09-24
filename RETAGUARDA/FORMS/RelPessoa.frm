VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClienteVendedor 
   Caption         =   "Clientes Por Vendedor"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RelPessoa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Data Compra"
      Height          =   975
      Left            =   50
      TabIndex        =   10
      Top             =   1800
      Width           =   7815
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Final:"
         Height          =   240
         Left            =   3435
         TabIndex        =   12
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inicial:"
         Height          =   240
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   50
      TabIndex        =   5
      Top             =   840
      Width           =   7815
      Begin VB.ComboBox cmbVendAux 
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
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbVend 
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
         Left            =   4095
         TabIndex        =   1
         Top             =   360
         Width           =   3585
      End
      Begin VB.ComboBox cmbEstabAUX 
         BackColor       =   &H00C0FFFF&
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
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.ComboBox cmbEstab 
         Height          =   360
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor(a):"
         Height          =   240
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Estab.:"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   630
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
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
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5040
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
               Picture         =   "RelPessoa.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RelPessoa.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RelPessoa.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RelPessoa.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RelPessoa.frx":9EFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RelPessoa.frx":BEDD
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmClienteVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If TIPO_PESSOA_CADASTRO = "CLIENTE" Then
      Me.Caption = "Clientes Por Estabelecimento"
   End If

   cmbEstabAUX.Clear
   cmbEstab.Clear
   cmbEstab.AddItem "Todos"
   cmbEstabAUX.AddItem ""

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select ESTABELECIMENTO_id,descricao from ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " where EMPRESA_id = " & EMPRESA_ID_N
   SQL = SQL & " order by DESCRICAO"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbEstab.AddItem Trim(TabTemp!DESCRICAO) & "-" & Trim(TabTemp.Fields("ESTABELECIMENTO_id").Value)
      cmbEstabAUX.AddItem Trim(TabTemp.Fields("ESTABELECIMENTO_id").Value)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   cmbEstabAUX.Text = ESTABELECIMENTO_ID_N
   cmbEstab.Text = "" & TRAZ_ESTABELECIMENTO(cmbEstabAUX.Text)

   cmbVend.Clear
   cmbVendAux.Clear
   cmbVend.AddItem "Todos"
   cmbVendAux.AddItem ""

   SQL = "select vendedor_id,descricao from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by descricao "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbVend.AddItem Trim(TabTemp!DESCRICAO) & " - " & Trim(TabTemp!VENDEDOR_ID)
      cmbVendAux.AddItem Trim(TabTemp!VENDEDOR_ID)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         cmbEstab.Text = ""
         cmbEstabAUX.Text = ""
         cmbVend.Text = ""
         cmbVendAux.Text = ""
         txtDtIni.PromptInclude = False
         txtDtIni.Text = ""
         txtDtFim.PromptInclude = False
         txtDtFim.Text = ""
         txtDtIni.SetFocus
      Case "voltar"
         TIPO_PESSOA_CADASTRO = ""
         CRITERIO_A = ""
         SQL = ""
         SqL2 = ""
         SQL3 = ""
         Unload Me
      Case "print"
         MONTA_REL
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmbvend_Click"
End Sub

Private Sub cmbestab_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.ListIndex = cmbEstab.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmbestab_Click"
End Sub

Private Sub txtDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtDTINI_GotFocus"
End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtDTINI_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtDTfim_GotFocus"
End Sub

Private Sub TXTDTFIM_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtDTfim_KeyPress"
End Sub

Sub MONTA_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{vwRelCliente.pedido_id} > 0"

   If Trim(cmbEstabAUX.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {vwRelCliente.estabelecimento_id} = " & Trim(cmbEstabAUX.Text)
   If Trim(cmbVendAux.Text) <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {vwRelCliente.vendedor_id} = " & Trim(cmbVendAux.Text)

   If Trim(txtDtIni.Text) <> "" Then
      DATA_INI = DMA(txtDtIni.Text, "i")
      DATA_FIM = DMA(txtDtFim.Text, "f")

      FORMULA_REL = FORMULA_REL & " and {vwRelCliente.DT_req} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
      FORMULA_REL = FORMULA_REL & " and {vwRelCliente.DT_req} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"
   End If

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_Cli_Estab.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "MONTA_REL"
End Sub
