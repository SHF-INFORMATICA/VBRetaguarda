VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCLIENTECONSULTA 
   Caption         =   "Consulta Clientes"
   ClientHeight    =   7530
   ClientLeft      =   1410
   ClientTop       =   2130
   ClientWidth     =   12675
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "consultacliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   12675
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   -120
      TabIndex        =   7
      Top             =   720
      Width           =   12735
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   21
         Top             =   320
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   " Vendas por Período "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1215
         Left            =   9240
         TabIndex        =   18
         Top             =   120
         Width           =   2775
         Begin MSMask.MaskEdBox txtDtIni 
            Height          =   315
            Left            =   1200
            TabIndex        =   0
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox txtDtFim 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Data Final:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   20
            Top             =   765
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data Inicial:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   19
            Top             =   405
            Width           =   1020
         End
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
         Left            =   1320
         TabIndex        =   15
         Top             =   1920
         Width           =   5175
      End
      Begin VB.ComboBox cmbUF 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7680
         TabIndex        =   5
         Top             =   1420
         Width           =   1215
      End
      Begin VB.ComboBox cmbCidade 
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
         Left            =   1320
         TabIndex        =   6
         Top             =   1400
         Width           =   5175
      End
      Begin VB.TextBox txtFone 
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
         Left            =   4800
         MaxLength       =   100
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtNome 
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
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   2
         Top             =   320
         Width           =   5175
      End
      Begin MSMask.MaskEdBox txtCep 
         Height          =   375
         Left            =   7440
         TabIndex        =   4
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtaux 
         Height          =   450
         Left            =   1440
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   794
         _Version        =   393216
         BackColor       =   0
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
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   6705
         TabIndex        =   22
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblVend 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   450
         TabIndex        =   16
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
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
         Left            =   7335
         TabIndex        =   13
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
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
         Left            =   675
         TabIndex        =   12
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:"
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
         Left            =   4275
         TabIndex        =   11
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
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
         Left            =   6930
         TabIndex        =   10
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CGC/CPF:"
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
         Left            =   405
         TabIndex        =   9
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblNome 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
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
         Left            =   765
         TabIndex        =   8
         Top             =   360
         Width           =   570
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultacliente.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultacliente.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultacliente.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultacliente.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultacliente.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultacliente.frx":707E
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
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   1270
      ButtonWidth     =   2858
      ButtonHeight    =   1111
      Appearance      =   1
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
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
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
         Left            =   5520
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5880
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "consultacliente.frx":739A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "consultacliente.frx":87C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "consultacliente.frx":9851
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "consultacliente.frx":A806
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "consultacliente.frx":B911
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView LISTA 
      Height          =   4260
      Left            =   0
      TabIndex        =   17
      Top             =   3240
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7514
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      ImageList       =   "ILTw"
      BorderStyle     =   1
      Appearance      =   1
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
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   12675
      DesignHeight    =   7530
   End
End
Attribute VB_Name = "frmCLIENTECONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim CRITERIO_A_Relatorio As String, QTD_COTAS As Long

Private Sub Form_Load()
   Call CentralizaJanela(frmCLIENTECONSULTA)

   MONTA_UF
   MOSTRA_VENDEDOR

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
End Sub

Private Sub Form_Resize()
   INDR_PRI = True

   MOSTRA_RODAPE "ESC - Sair", "", "", "", ""

   cmbSTATUS.Clear
   cmbSTATUS.AddItem "Ativo"
   cmbSTATUS.AddItem "Inativo"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "sair"
         MOSTRA_RODAPE "Aguarde ...", "", "", "", ""
         Unload Me
      Case "print"
         IMPRIME_CLIENTE
      Case "consultar"
         MOSTRA_RODAPE "Aguarde, Pesquisando ...", "", "", "", ""

         HORA_INI = Time
            MONTA_CONSULTA
         HORA_FIM = Time

         MOSTRA_RODAPE "ESC - Sair", "Duplo click para selecionar", "Tempo consulta = " & Format((HORA_FIM - HORA_INI), "hh:mm:ss"), "Registros Encontrados = " & NUMR_CONSULTA_N, ""
      Case "limpar"
         LIMPA_CLI
         MOSTRA_RODAPE "Aguarde ...", "ESC - Sair", "", "", ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCpf_Click()
   txtNome.SetFocus
End Sub

Sub MOSTRA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   cmbVend.Clear

   If TabEQUIPE.State = 1 Then _
      TabEQUIPE.Close

   SQL = "select * from vwVendedor WITH (NOLOCK) "
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " order by descricao"
   TabEQUIPE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabVENDEDOR.EOF
      cmbVend.AddItem TabEQUIPE!DESCRICAO & "-" & TabEQUIPE!VENDEDOR_ID
      cmbVend.ItemData(cmbVend.ListCount - 1) = TabEQUIPE!VENDEDOR_ID
      TabEQUIPE.MoveNext
   Wend
   If TabEQUIPE.State = 1 Then _
      TabEQUIPE.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VENDEDOR"
End Sub

Private Sub txtCep_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe o CEP", "", "", ""
End Sub

Private Sub TXTCNPJCPF_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe número do CGC/CPF", "", "", ""

   txtCNPJCPF.Mask = "##############"
End Sub

Private Sub TXTDTINI_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe data inicial", "", "", ""

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
End Sub

Private Sub TXTDTFIM_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe data final"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   
   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If
End Sub

Private Sub txtFone_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o telefone"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub txtNome_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o nome"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub
Private Sub cmbCidade_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Selecione uma cidade"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub cmbUF_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Selecione um estado"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub cmbUF_Click()
'On Error GoTo ERRO_TRATA

   If cmbUF.Text <> "" Then
      If cmbCidade.Enabled = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select distinct(cidade),uf from CEP "
         SQL = SQL & " where uf='" & cmbUF.Text & "'"
         SQL = SQL & " order by cidade"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            cmbCidade.AddItem TabTemp!CIDADE
            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         cmbCidade.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbUF_Click"
End Sub

Private Sub MONTA_UF()
'On Error GoTo ERRO_TRATA

   cmbUF.Clear
   cmbCidade.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select distinct(uf) from CEP "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbUF.AddItem TabTemp!UF
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_UF"
End Sub

Private Sub LIMPA_CLI()
'On Error GoTo ERRO_TRATA

   cmbVend.Clear
   cmbSTATUS.Text = ""
   cmbVend.Clear
   LISTA.Nodes.Clear
   txtNome.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCep.Text = ""
   txtFone.Text = ""
   cmbCidade.Text = ""
   cmbCidade.Clear
   cmbUF.Text = ""
   cmbUF.Clear
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_CLI"
End Sub
'========
Private Sub MONTA_CONSULTA()
'On Error GoTo ERRO_TRATA

   CRITERIO_A_Relatorio = ""
   If IsDate(txtDtIni.Text) Then
      If IsDate(txtDtFim.Text) Then
         SQL = "select distinct(c.cgccpf), c.nome,w.cgccpf,w.dt_req from CLIENTE c, PEDIDO w "
         Else: SQL = "select distinct(c.cgccpf), c.nome from CLIENTE c "
      End If
      Else: SQL = "select distinct(c.cgccpf), c.nome from CLIENTE c "
   End If

   SQL = SQL & " where c.estabelecimento_id <> 0 "

   If Trim(txtNome.Text) <> "" Then
      SQL = SQL & " and c.nome like '" & txtNome.Text & "%" & "'"
      'RELATÓRIO
      CRITERIO_A_Relatorio = "({CLIENTE.nome} like '" & UCase(txtNome.Text) & "%" & "'"
      CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " or {CLIENTE.nome} like '" & LCase(txtNome.Text) & "%" & "'"
      CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " or {CLIENTE.nome} like '" & UCase(Left(txtNome.Text, 1)) & Mid(LCase(txtNome.Text), 2, Len(txtNome.Text)) & "%" & "'" & ")"
   End If

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" Then
      CRITERIO_A = Chr$(39) & txtCNPJCPF.Text & "%" & Chr(39)
      SQL = SQL & " and c.CGCCPF like " & CRITERIO_A
      'RELATÓRIO
      If CRITERIO_A_Relatorio = "" Then
         CRITERIO_A_Relatorio = "{CLIENTE.cgccpf} like " & CRITERIO_A
         Else: CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {CLIENTE.cgccpf} like " & CRITERIO_A
      End If
   End If

   If txtFone.Text <> "" Then
      CRITERIO_A = Chr$(39) & txtFone.Text & "%" & Chr(39)
      SQL = "select distinct(c.cgccpf), c.nome from CLIENTE c, FONE f "
      SQL = SQL & " where c.pessoa_id = f.pessoa_id "
      SQL = SQL & "and f.numero like " & CRITERIO_A
      If CRITERIO_A_Relatorio = "" Then
         CRITERIO_A_Relatorio = "{FONE.numero} like " & txtFone.Text
         Else: CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {FONE.numero} like " & txtFone.Text
      End If
   End If

   If txtCep.Text <> "" Then
      CRITERIO_A = Chr$(39) & txtCep.Text & "%" & Chr(39)
      If CRITERIO_A_Relatorio = "" Then
         CRITERIO_A_Relatorio = "{CEP.Cep_ID} like " & CRITERIO_A
         Else: CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {CEP.Cep_ID} like " & CRITERIO_A
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
      SQL = SQL & " where c.pessoa_id = e.pessoa_id "
      SQL = SQL & " and p.Cep_ID = e.Cep_ID "
      SQL = SQL & "and p.Cep_ID like " & CRITERIO_A
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
         SQL = SQL & " where c.pessoa_id = e.pessoa_id "
         SQL = SQL & " and p.Cep_ID = e.Cep_ID "
         SQL = SQL & "and p.Cep_ID like " & CRITERIO_A
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
            SQL = SQL & " where c.pessoa_id = e.pessoa_id "
            SQL = SQL & "and p.Cep_ID = e.Cep_ID "
            SQL = SQL & "and p.Cep_ID like " & CRITERIO_A
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   If cmbUF.Text <> "" Then
      If CRITERIO_A_Relatorio = "" Then
         CRITERIO_A_Relatorio = "{CEP.uf} = '" & cmbUF.Text & "'"
         Else: CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {CEP.uf} = '" & cmbUF.Text & "'"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
      SQL = SQL & " where c.pessoa_id = e.pessoa_id "
      SQL = SQL & "and p.Cep_ID = e.Cep_ID "
      SQL = SQL & "and p.uf='" & cmbUF.Text & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
         SQL = SQL & " where c.pessoa_id = e.pessoa_id "
         SQL = SQL & "and p.Cep_ID = e.Cep_ID "
         SQL = SQL & "and p.uf='" & cmbUF.Text & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
            SQL = SQL & " where c.pessoa_id = e.pessoa_id "
            SQL = SQL & "and p.Cep_ID=e.Cep_ID "
            SQL = SQL & "and p.uf='" & cmbUF.Text & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   If cmbCidade.Text <> "" Then
      CRITERIO_A = Chr$(39) & cmbCidade.Text & "%" & Chr(39)
      If CRITERIO_A_Relatorio = "" Then
         CRITERIO_A_Relatorio = "{CEP.cidade} like " & CRITERIO_A
         Else: CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {CEP.cidade} like " & CRITERIO_A
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
      SQL = SQL & " where c.pessoa_id = e.pessoa_id "
      SQL = SQL & "and p.Cep_ID=e.Cep_ID "
      SQL = SQL & "and p.cidade like " & CRITERIO_A
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
         SQL = SQL & " where c.pessoa_id = e.pessoa_id "
         SQL = SQL & "and p.Cep_ID = e.Cep_ID "
         SQL = SQL & "and p.cidade like " & CRITERIO_A
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select distinct(CGCCPF),nome from CLIENTE c, CEP p, ENDERECO e "
            SQL = SQL & " where c.pessoa_id = e.pessoa_id "
            SQL = SQL & "and p.Cep_ID = e.Cep_ID "
            SQL = SQL & "and p.cidade like " & CRITERIO_A
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   If cmbVend.Text <> "" Then
      If CRITERIO_A_Relatorio = "" Then
         CRITERIO_A_Relatorio = "{PEDIDO.vendedor_id} = " & cmbVend.ItemData(cmbVend.ListIndex)
         Else: CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {PEDIDO.vendedor_id} = " & cmbVend.ItemData(cmbVend.ListIndex)
      End If

      SQL = "select distinct(c.cgccpf),c.nome,w.vendedor_id,v.vendedor_id from CLIENTE c, VENDEDOR v, PEDIDO w "
      SQL = SQL & " where w.vendedor_id=v.vendedor_id "
      SQL = SQL & " and w.cgccpf=c.cgccpf "
      SQL = SQL & " and w.vendedor_id = " & cmbVend.ItemData(cmbVend.ListIndex)
      SQL = SQL & " and v.vendedor_id = " & cmbVend.ItemData(cmbVend.ListIndex)
   End If

   If cmbSTATUS.Text <> "" Then
      SQL = SQL & " and c.status='" & Left(cmbSTATUS.Text, 1) & "'"
         If CRITERIO_A_Relatorio = "" Then
            CRITERIO_A_Relatorio = "{CLIENTE.status} = '" & Left(cmbSTATUS.Text, 1) & "'"
            Else: CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {CLIENTE.status} = '" & Left(cmbSTATUS.Text, 1) & "'"
         End If
   End If

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and c.cgccpf=w.cgccpf "
      MONTA_DATAS
      If CRITERIO_A_Relatorio <> "" Then
         CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {PEDIDO.dt_req} >= DATE (" & Year(txtDtIni.Text) & "," & Month(txtDtIni.Text) & "," & Day(txtDtIni.Text) & ") "
         CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {PEDIDO.dt_req} <= DATE (" & Year(txtDtFim.Text) & "," & Month(txtDtFim.Text) & "," & Day(txtDtFim.Text) & ")"
         Else
            CRITERIO_A_Relatorio = "{PEDIDO.dt_req} >= DATE (" & Year(txtDtIni.Text) & "," & Month(txtDtIni.Text) & "," & Day(txtDtIni.Text) & ") "
            CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {PEDIDO.dt_req} <= DATE (" & Year(txtDtFim.Text) & "," & Month(txtDtFim.Text) & "," & Day(txtDtFim.Text) & ")"
      End If
      CRITERIO_A_Relatorio = CRITERIO_A_Relatorio & " and {CLIENTE.cgccpf} = {PEDIDO.cgccpf} "
   End If

   SETA_GRID 'MOSTRA DADOS TELA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_CONSULTA"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Cont(2)

   HORA_INI = Time
   LISTA.Nodes.Clear
   NUMR_SEQ_N = 0
   NUMR_CONSULTA_N = 0
   CONT_N = 0
   Cont(1) = "Endereço"
   QTD_COTAS = 0
   CRITERIO_A = ""

   If TabCliente.State = 1 Then _
      TabCliente.Close

   SQL = SQL & " order by c.nome "

   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCliente.EOF
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      txtaux.PromptInclude = False
      If Len(Trim(TabCliente(0))) <= 11 Then
         txtaux.Mask = "###.###.###-##"
         txtaux.Text = TabCliente(0)
         txtaux.PromptInclude = True
         Set Nodx = LISTA.Nodes.Add(, , Cont(1) & QTD_COTAS, txtaux.Text & " - " & TabCliente!NOME)
         Else
            txtaux.Mask = "##.###.###/####-##"
            txtaux.Text = TabCliente(0)
            txtaux.PromptInclude = True
            Set Nodx = LISTA.Nodes.Add(, , Cont(1) & QTD_COTAS, txtaux.Text & " - " & TabCliente!NOME)
      End If
      'ENDEREÇO
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      SQL = "select * from ENDERECO "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabEndereco.EOF Then
         CRITERIO_A = "Rua " & Replace("" & tabEndereco!Rua, ",", ".") & " " & Replace("" & tabEndereco!Complemento, ",", ".") & " Bairro " & Replace("" & tabEndereco!Bairro, ",", ".")
         If Not IsNull(tabEndereco!CEP_ID) Then
            If tabEndereco!CEP_ID <> "" Then
               If TabCEP.State = 1 Then _
                  TabCEP.Close

               SQL = "select * from CEP "
               SQL = SQL & " where cep_ID = " & tabEndereco!CEP_ID
               TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCEP.EOF Then _
                  CRITERIO_A = CRITERIO_A & " Cidade " & TabCEP!CIDADE & "-" & TabCEP!UF

               If TabCEP.State = 1 Then _
                  TabCEP.Close
            End If
         End If
         Set Nodx = LISTA.Nodes.Add(Cont(1) & QTD_COTAS, tvwChild, "endereco" & NUMR_SEQ_N, CRITERIO_A)
         NUMR_SEQ_N = NUMR_SEQ_N + 1
      End If
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

'FONES
      If TabFone.State = 1 Then _
         TabFone.Close

      SQL = "select * from FONE "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabFone.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFone.EOF Then _
         Set Nodx = LISTA.Nodes.Add(Cont(1) & QTD_COTAS, tvwChild, "fone" & CONT_N, "Telefone(s)")
      While Not TabFone.EOF
         Set Nodx = LISTA.Nodes.Add("fone" & CONT_N, tvwChild, , TabFone!local & " " & TabFone!DDD & " " & TabFone!Numero)
         TabFone.MoveNext
      Wend
      If TabFone.State = 1 Then _
         TabFone.Close

'VENDAS
      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      SQL = "select * from PEDIDO "
      SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and cgccpf = '" & TabCliente(0) & "'"

      If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then _
         MONTA_DATAS

      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCabeca.EOF Then _
         Set Nodx = LISTA.Nodes.Add(Cont(1) & QTD_COTAS, tvwChild, "contra" & CONT_N, "Requisições")
      While Not TabCabeca.EOF
         CRITERIO_A = ""
         NOME_A = ""
         If Not IsNull(TabCabeca!STATUS) Then
'1=ORÇAMENTO;2=EMITIDA;3=EMITIDA COM NOTA;4=EMITIDA COM CUPOM;9=CANCELADO
            If TabCabeca!STATUS = 1 Then _
               CRITERIO_A = "ORÇAMENTO"
            If TabCabeca!STATUS = 2 Then _
               CRITERIO_A = "Requisição Emitida"
            If TabCabeca!STATUS = 3 Then _
               CRITERIO_A = "Requisição Emitida com Nota"
            If TabCabeca!STATUS = 4 Then _
               CRITERIO_A = "Requisição Emitida com Cupom"
            If TabCabeca!STATUS = 9 Then _
               CRITERIO_A = "Requisição Cancelada"
         End If
         SQL = "select descricao from vwVendedor WITH (NOLOCK) "
         SQL = SQL & " where vendedor_id = " & TabCabeca!VENDEDOR_ID
         If TabVENDEDOR.State = 1 Then TabVENDEDOR.Close
         TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabVENDEDOR.EOF Then _
            NOME_A = TabVENDEDOR!DESCRICAO
         TabVENDEDOR.Close
         Set Nodx = LISTA.Nodes.Add("contra" & CONT_N, tvwChild, , TabCabeca!PEDIDO_ID & " , Data Venda: " & TabCabeca!DT_REQ & " , vendedor_id = " & NOME_A & " , Status: " & CRITERIO_A & SqL2)
         SqL2 = " "
         TabCabeca.MoveNext
      Wend
      If TabCabeca.State = 1 Then _
         TabCabeca.Close
      If TabFone.State = 1 Then _
         TabFone.Close

      QTD_COTAS = QTD_COTAS + 1
      CONT_N = CONT_N + 1
      TabCliente.MoveNext
      CRITERIO_A = ""
   Wend
   If TabCliente.State = 1 Then _
      TabCliente.Close

   LISTA.Refresh
   HORA_FIM = Time

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub MONTA_DATAS()
'On Error GoTo ERRO_TRATA

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_req >= '" & DMA(txtDtIni.Text, "i") & "'"
      SQL = SQL & " and dt_req <= '" & DMA(txtDtFim.Text, "f") & "'"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MONTA_DATAS"
End Sub
'==================================
Private Sub IMPRIME_CLIENTE()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Pesquisando, Aguarde ... "
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   FORMULA_REL = "{CLIENTE.nome} <> '' "

   If cmbVend.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " {CLIENTE.vendedor_id} = " & cmbVend.ItemData(cmbVend.ListIndex)

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_Cliente.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIME_CLIENTE"
End Sub
