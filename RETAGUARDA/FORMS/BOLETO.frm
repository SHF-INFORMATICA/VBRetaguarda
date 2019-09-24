VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRELBOLETO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de Boleto Bancário Avulso"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "BOLETO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      Height          =   240
      Left            =   7080
      TabIndex        =   42
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox txtdup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      MaxLength       =   6
      TabIndex        =   40
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton opthsbc 
      Caption         =   "Option2"
      Height          =   255
      Left            =   5640
      TabIndex        =   37
      Top             =   7920
      Width           =   255
   End
   Begin VB.OptionButton optbradesco 
      Caption         =   "Option1"
      Height          =   255
      Left            =   5640
      TabIndex        =   36
      Top             =   7560
      Width           =   255
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      Picture         =   "BOLETO.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAIRentrada 
      Caption         =   "&Sair"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "BOLETO.frx":6054
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLIMPARentrada 
      Caption         =   "Li&mpar"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      Picture         =   "BOLETO.frx":6496
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   -120
      TabIndex        =   26
      Top             =   5640
      Width           =   8775
      Begin VB.OptionButton optM 
         Caption         =   "Cobrança"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optC 
         Caption         =   "Comercial"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optR 
         Caption         =   "Residêncial"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtUF2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1200
         Width           =   600
      End
      Begin VB.TextBox txtCep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCidade2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtEnd2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   10
         Top             =   720
         Width           =   6495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   30
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cep:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   29
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Município:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   28
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   27
         Top             =   720
         Width           =   1350
      End
   End
   Begin VB.TextBox txtCli 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      MaxLength       =   70
      TabIndex        =   8
      Top             =   5040
      Width           =   5295
   End
   Begin VB.TextBox txtDtEmis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtDoc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtValorFatura 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      MaxLength       =   12
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtDtVenc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtLOCAL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      MaxLength       =   60
      TabIndex        =   6
      Text            =   "Pagável em qualquer banco até o vencimento"
      Top             =   2280
      Width           =   8415
   End
   Begin VB.TextBox txtINSTRUCAO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "BOLETO.frx":7160
      Top             =   3360
      Width           =   8415
   End
   Begin VB.TextBox txtTipoDoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      TabIndex        =   3
      Text            =   "N"
      Top             =   720
      Width           =   360
   End
   Begin MSMask.MaskEdBox txtCGCCPF 
      Height          =   420
      Left            =   5520
      TabIndex        =   9
      Top             =   5040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lbldup 
      AutoSize        =   -1  'True
      Caption         =   "Dup:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      TabIndex        =   41
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lbltipo1 
      Caption         =   "Tipo HSBC"
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
      Left            =   6000
      TabIndex        =   39
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label lbltipo 
      Caption         =   "Tipo Bradesco"
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
      Left            =   6000
      TabIndex        =   38
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "Tipo Boleto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Seqüência:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   34
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ/CPF"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6120
      TabIndex        =   25
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data da Emissão:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   23
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nº Documento:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   22
      Top             =   120
      Width           =   1950
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor Documento:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data Vencimento:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Local de Pagamento"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   2700
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Instruções:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   1650
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Aceite:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   17
      Top             =   720
      Width           =   1050
   End
End
Attribute VB_Name = "frmRELBOLETO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   MOSTRA_RODAPE "ESC - SAIR", "F10 - IMPRIMIR", "", "", ""

   LIMPA_TUDO
   If NUMR_REQ_N > 0 Then
      txtDoc.Text = NUMR_REQ_N
      txtSeq.Text = SQL3
      PROCURA_BOLETO
   End If

   Call CentralizaJanela(frmRELBOLETO)
   optbradesco.Value = True
End Sub

Private Sub cmdSAIRentrada_Click()
   Unload Me
End Sub

Private Sub cmdLIMPARentrada_Click()
   LIMPA_TUDO
   txtDoc.SetFocus
End Sub

Private Sub optbradesco_Click()
   optbradesco.Enabled = True
End Sub

Private Sub opthsbc_Click()
   opthsbc.Enabled = True
End Sub

Private Sub cmdPrint_Click()
   If txtDtVenc.Text = "" Then
      MsgBox "Data de vencimento inválida."
      Exit Sub
   End If
   If txtDtEmis.Text = "" Then
      MsgBox "Data de emissão inválida."
      Exit Sub
   End If
   If txtDoc.Text = "" Then
      MsgBox "Número documento inválido."
      Exit Sub
   End If
   If txtValorFatura.Text = "" Then
      MsgBox "Valor documento inválido."
      Exit Sub
   End If
   If txtLOCAL.Text = "" Then
      MsgBox "Local de pagamento inválido."
      Exit Sub
   End If

   GRAVA_BOLETO

   FORMULA_REL = "{BOLETO.numr_doc} = " & txtDoc.Text
   FORMULA_REL = FORMULA_REL & " and {BOLETO.seq} = " & txtSeq.Text

   If optbradesco.Value = True Then
      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "rel_boleto.rpt"
      frmRELATORIO10.Show 1
   End If

   If opthsbc.Value = True Then
      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "rel_boleto_hsbc.rpt"
      frmRELATORIO10.Show 1
   End If
End Sub

Private Sub TXTCGCCPF_GotFocus()
   txtCGCCPF.PromptInclude = False

   If txtCGCCPF.Text <> "" Then _
      CRITERIO = txtCGCCPF.Text
      txtCGCCPF.Mask = "##############"

   txtCGCCPF.PromptInclude = True
End Sub

Private Sub txtCli_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYCLIENTE.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCGCCPF.PromptInclude = False
            txtCGCCPF.Text = CNPJCPF_A

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select nome from CLIENTE "
            SQL = SQL & " where cgccpf = '" & CNPJCPF_A & "'"
            SQL = SQL & " and status = 'A'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then _
               txtCli.Text = Trim(TabTemp.Fields(0).Value)

            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
   End Select
End Sub

Private Sub txtDTVENC_GotFocus()
   If Trim(txtDtVenc.Text) = "" Then
      txtDtVenc.Text = Date
   End If
End Sub

Private Sub txtDtEmis_GotFocus()
   If Trim(txtDtEmis.Text) = "" Then
      txtDtEmis.Text = Date
   End If
End Sub

Private Sub txtValorFatura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtValorFatura_LostFocus()
On Error GoTo ERRO_TRATA

   Dim VLR_JUROS As Double
   Dim PERC_JUROS As Double
   Dim PERC_FORMAT As Double

   VLR_JUROS = 0
   PERC_JUROS = 0
   PERC_FORMAT = 0

   If txtValorFatura.Text <> "" Then
      txtValorFatura.Text = Format(txtValorFatura.Text, strFormatacao2Digitos)

      VLR_JUROS = txtValorFatura.Text
      PERC_JUROS = "5,4" / 100

      VLR_JUROS = VLR_JUROS * PERC_JUROS

      CRITERIO = "COBRAR MORA DIÁRIA DE R$ " & Format(VLR_JUROS, strFormatacao2Digitos)
      CRITERIO = CRITERIO & " POR DIA DE ATRAZO PROTESTAR APÓS 5 DIAS DE VENCIMENTO "

      txtINSTRUCAO.Text = CRITERIO
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorFatura_LostFocus"
End Sub

Private Sub optC_Click()
   CHECA_CGCCPF
End Sub

Private Sub optM_Click()
   CHECA_CGCCPF
End Sub

Private Sub optR_Click()
   CHECA_CGCCPF
End Sub

Private Sub txtCGCCPF_LostFocus()
   txtCGCCPF.PromptInclude = False
   If txtCGCCPF.Text <> "" Then
      CRITERIO = txtCGCCPF.Text
      txtCGCCPF.Mask = "##############"

      If Len(txtCGCCPF.Text) <= 11 Then _
         txtCGCCPF.Mask = "###.###.###-##"

      If Len(txtCGCCPF.Text) > 11 Then _
         txtCGCCPF.Mask = "##.###.###/####-##"

      txtCGCCPF.Text = CRITERIO
   End If

   CNPJCPF_A = txtCGCCPF.Text

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select nome from CLIENTE "
   SQL = SQL & " where cgccpf = '" & CNPJCPF_A & "'"
   SQL = SQL & " and status = 'A'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      txtCli.Text = Trim(TabTemp.Fields(0).Value)

   If TabTemp.State = 1 Then _
      TabTemp.Close

   txtCGCCPF.PromptInclude = True
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtDoc.Text <> "" Then
         CRITERIO = txtDoc.Text
         txtSeq.SetFocus
      End If
   End If
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtDoc.Text <> "" Then
         CRITERIO = txtDoc.Text
         NUMR_SEQ_N = 0 & txtSeq.Text
         PROCURA_BOLETO
      End If
   End If
End Sub
'==================================
Private Sub GRAVA_BOLETO()
On Error GoTo ERRO_TRATA

   Dim VLR_JUROS As Double
   Dim PERC_JUROS As Double
   Dim PERC_FORMAT As Double
   VLR_JUROS = 0
   PERC_JUROS = 0
   PERC_FORMAT = 0

   VLR_JUROS = (Format(txtValorFatura.Text, strFormatacao2Digitos) / 100)
   PERC_JUROS = ((VLR_JUROS * 7) / 30)
   PERC_FORMAT = Format(PERC_JUROS, strFormatacao2Digitos)

   CRITERIO = "COBRAR MORA DIÁRIA DE R$ "
   CRITERIO = " " & CRITERIO & PERC_FORMAT
   CRITERIO = CRITERIO & "POR DIA DE ATRAZO PROTESTAR APÓS 5 DIAS DE VENCIMENTO "

   CRITERIO = Trim(txtINSTRUCAO.Text)

   txtCGCCPF.PromptInclude = True

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from BOLETO "
   SQL = SQL & " where numr_doc = " & txtDoc.Text
   SQL = SQL & " and seq = " & txtSeq.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "update BOLETO set "
      SQL = SQL & "EMPRESA_ID = " & EMPRESA_ID_N                  '[EMPRESA_ID]
      SQL = SQL & ",NUMR_DOC = '" & Trim(txtDoc.Text) & "'"       '[NUMR_DOC]
      SQL = SQL & ",SEQ = " & Trim(txtSeq.Text)                   '[SEQ]
      SQL = SQL & ",NUMR_DP = '" & Trim(txtdup.Text) & "'"        '[NUMR_DP]
      SQL = SQL & ",CGCCPF = '" & Trim(txtCGCCPF.Text) & "'"      '[CGCCPF]
      SQL = SQL & ",LOCAL_PAGTO = '" & Trim(txtLOCAL.Text) & "'"  '[LOCAL_PAGTO]
      SQL = SQL & ",DT_VENC = '" & DMA(txtDtVenc.Text) & "'"      '[DT_VENC]
      SQL = SQL & ",DT_DOC = '" & DMA(Date) & "'"                 '[DT_DOC]
      SQL = SQL & ",VALOR_DOC = " & Str(txtValorFatura.Text)     '[VALOR_DOC]
      SQL = SQL & ",VALOR_DESCONTO = 0"                           '[VALOR_DESCONTO]
      SQL = SQL & ",VALOR_COBRADO = 0"                            '[VALOR_COBRADO]
      SQL = SQL & ",INSTRUCAO = '" & CRITERIO & "'"               '[INSTRUCAO]
      SQL = SQL & ",TIPO_DOC = '" & Trim(txtTipoDoc.Text) & "'"   '[TIPO_DOC]
      SQL = SQL & ",CLIENTE = '" & Trim(txtCli.Text) & "'"        '[CLIENTE]
      SQL = SQL & ",ENDERECO = '" & Trim(txtEnd2.Text) & "'"      '[ENDERECO]
      SQL = SQL & ",UF = '" & Trim(txtUF2.Text) & "'"             '[UF]
      SQL = SQL & ",CIDADE = '" & Trim(txtCidade2.Text) & "'"     '[CIDADE]
      SQL = SQL & ",CEP = '" & Trim(txtCep.Text) & "'"            '[CEP]
      SQL = SQL & " where numr_doc = " & txtDoc.Text
      SQL = SQL & " and seq = " & txtSeq.Text
      Else
         SQL = "insert into BOLETO values( "
         SQL = SQL & EMPRESA_ID_N                        '[EMPRESA_ID]
         SQL = SQL & ",'" & Trim(txtDoc.Text) & "'"      '[NUMR_DOC]
         SQL = SQL & "," & Trim(txtSeq.Text)             '[SEQ]
         SQL = SQL & ",'" & Trim(txtdup.Text) & "'"      '[NUMR_DP]
         SQL = SQL & ",'" & Trim(txtCGCCPF.Text) & "'"   '[CGCCPF]
         SQL = SQL & ",'" & Trim(txtLOCAL.Text) & "'"    '[LOCAL_PAGTO]
         SQL = SQL & ",'" & DMA(txtDtVenc.Text) & "'"    '[DT_VENC]
         SQL = SQL & ",'" & DMA(Date) & "'"              '[DT_DOC]
         SQL = SQL & "," & Str(txtValorFatura.Text)     '[VALOR_DOC]
         SQL = SQL & ",0"                                '[VALOR_DESCONTO]
         SQL = SQL & ",0"                                '[VALOR_COBRADO]
         SQL = SQL & ",'" & CRITERIO & "'"               '[INSTRUCAO]
         SQL = SQL & ",'" & Trim(txtTipoDoc.Text) & "'"  '[TIPO_DOC]
         SQL = SQL & ",'" & Trim(txtCli.Text) & "'"      '[CLIENTE]
         SQL = SQL & ",'" & Trim(txtEnd2.Text) & "'"     '[ENDERECO]
         SQL = SQL & ",'" & Trim(txtUF2.Text) & "'"      '[UF]
         SQL = SQL & ",'" & Trim(txtCidade2.Text) & "'"  '[CIDADE]
         SQL = SQL & ",'" & Trim(txtCep.Text) & "'"      '[CEP]
         SQL = SQL & " )"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_BOLETO"
End Sub

Private Sub LIMPA_TUDO()
   txtDoc.Text = ""
   txtSeq.Text = ""
   txtValorFatura.Text = ""
   txtDtVenc.Text = ""
   txtDtEmis.Text = ""
   txtCli.Text = ""
   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = ""
   optR.Value = False
   optM.Value = False
   optC.Value = False
   LIMPA_ENDERECO
End Sub

Private Sub LIMPA_ENDERECO()
   txtCep.Text = ""
   txtCidade2.Text = ""
   txtUF2.Text = ""
   txtEnd2.Text = ""
End Sub

Private Sub PROCURA_BOLETO()
On Error GoTo ERRO_TRATA

   If txtDoc.Text <> "" Then
      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select * from BOLETO "
      SQL = SQL & " where numr_doc = " & txtDoc.Text
      SQL = SQL & " and seq = " & txtSeq.Text
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then
         txtCGCCPF.PromptInclude = False
         txtCGCCPF.Text = TabAUX!CGCCPF
         txtCGCCPF.PromptInclude = True
         txtDoc.Text = TabAUX!Numr_doc
         txtLOCAL.Text = TabAUX!LOCAL_PAGTO
         txtDtVenc.Text = TabAUX!Dt_Venc
         txtDtEmis.Text = TabAUX!DT_DOC
         txtValorFatura.Text = Format(TabAUX!VALOR_doc, strFormatacao2Digitos)
         txtINSTRUCAO.Text = TabAUX!INSTRUCAO
         txtTipoDoc.Text = TabAUX!TIPO_DOC
         txtCli.Text = TabAUX!Cliente & ""
         txtEnd2.Text = TabAUX!Endereco & ""
         txtUF2.Text = TabAUX!UF & ""
         txtCep.Text = TabAUX!CEP & ""
         txtCidade2.Text = TabAUX!Cidade & ""
      End If
      If TabAUX.State = 1 Then _
         TabAUX.Close
   End If

   txtDtEmis.Text = Format(Date, "dd/mm/yyyy")

   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   SQL = "select * from LANCAMENTO l, ITEMLANCAMENTO i "

   SQL = SQL & " where l.numr_doc = " & txtDoc.Text
   SQL = SQL & " and l.numr_doc = i.numr_doc "
   SQL = SQL & " and i.seq = " & txtSeq.Text
   SQL = SQL & " and l.estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Not TabLancamento.EOF Then
      txtValorFatura.Text = Format(TabLancamento!Valor_Item, strFormatacao2Digitos)
      txtDtVenc.Text = Format(TabLancamento!DT_VENCIMENTO, "dd/mm/yyyy")

      If Not IsNull(TabLancamento!NUMR_DP) Then _
         txtdup.Text = TabLancamento!NUMR_DP

      txtSeq.Text = TabLancamento!SEQ

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from CLIENTE "
      SQL = SQL & " where cgccpf='" & TabLancamento!prop & "'"
      SQL = SQL & " and status = 'A'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         txtCli.Text = TabCliente!NOME & ""
         txtCGCCPF.PromptInclude = False
            txtCGCCPF.Text = TabCliente!CGCCPF
         txtCGCCPF.PromptInclude = True

         If TabAUX.State = 1 Then _
            TabAUX.Close

         'COBRANÇA
         SQL = "select * from ENDERECO "
         SQL = SQL & " where prop='" & TabLancamento!prop & "'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            txtEnd2.Text = TabAUX!Rua & " , " & TabAUX!Bairro & " , " & TabAUX!Complemento
            If Not IsNull(TabAUX!CEP) Then
               If TabCEP.State = 1 Then _
                  TabCEP.Close

               SQL = "select * from CEP "
               SQL = SQL & " where cep = '" & TabAUX!CEP & "'"
               TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabCEP.EOF Then
                  txtCidade2.Text = TabCEP!Cidade
                  txtCep.Text = TabCEP!CEP
                  txtUF2.Text = TabCEP!UF
               End If
               If TabCEP.State = 1 Then _
                  TabCEP.Close
            End If
         End If

         If TabAUX.State = 1 Then _
            TabAUX.Close

         Else
            MsgBox "Cliente não cadastrado, verifique !!! " & TabLancamento!prop
            Exit Sub
      End If 'cliente

      If TabCliente.State = 1 Then _
         TabCliente.Close

      If TabTemp.State = 1 Then _
         TabEmpresa.Close

      SQL = "select instrucao_boleto from ESTABELECIMENTO"
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp!instrucao_boleto) Then _
            txtINSTRUCAO.Text = TabTemp!instrucao_boleto

      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   If TabLancamento.State = 1 Then _
      TabLancamento.Close

   CRITERIO = txtCGCCPF.Text
   txtCGCCPF.Mask = "##############"

   If Len(txtCGCCPF.Text) <= 11 Then _
      txtCGCCPF.Mask = "###.###.###-##"

   If Len(txtCGCCPF.Text) > 11 Then _
      txtCGCCPF.Mask = "##.###.###/####-##"

   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = CRITERIO

   txtCGCCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_BOLETO"
End Sub

Private Sub CHECA_CGCCPF()
On Error GoTo ERRO_TRATA

   Dim Tipo_Endereco As String

   txtCGCCPF.PromptInclude = False
   CRITERIO = Replace(txtCGCCPF.Text, ".", "")
   CRITERIO = Replace(CRITERIO, "/", "")
   CRITERIO = Replace(CRITERIO, "-", "")
   CRITERIO = Trim(CRITERIO)
   If Trim(CRITERIO) <> "" Then
      If optR.Value = True Then _
         Tipo_Endereco = "R"
      If optC.Value = True Then _
         Tipo_Endereco = "C"
      If optM.Value = True Then _
         Tipo_Endereco = "B"

      SP_PROCURA_ENDEREÇONFE Trim(CRITERIO), Tipo_Endereco, Trim(txtCep.Text), 0

      If Not tabEndereco.EOF Then
         LIMPA_ENDERECO

         txtEnd2.Text = tabEndereco!Rua & " , " & tabEndereco!Bairro & " , " & tabEndereco!Complemento
         If Not IsNull(tabEndereco.Fields("Cep").Value) Then

            If TabCEP.State = 1 Then _
               TabCEP.Close

            SQL = "select * from CEP "
            SQL = SQL & " where cep = '" & tabEndereco.Fields("Cep").Value & "'"

            TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

            If Not TabCEP.EOF Then
               txtCidade2.Text = TabCEP!Cidade
               txtCep.Text = TabCEP!CEP
               txtUF2.Text = TabCEP!UF
            End If

            If TabCEP.State = 1 Then _
               TabCEP.Close
         End If
         Else: MsgBox "Não existe endereço residêncial cadastrado para esse cliente."
      End If
      If tabEndereco.State = 1 Then _
         tabEndereco.Close
   End If
   txtCGCCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_CGCCPF"
End Sub

