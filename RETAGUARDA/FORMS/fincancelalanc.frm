VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFINCANCELALANC 
   Caption         =   "Cancelar Lançamentos Contas à Pagar/Receber"
   ClientHeight    =   7230
   ClientLeft      =   2295
   ClientTop       =   2250
   ClientWidth     =   10965
   Icon            =   "fincancelalanc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10965
   Begin VB.ComboBox cmbStatusLanc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5280
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtLanc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   120
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbAuxStatusLanc 
      BackColor       =   &H80000000&
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
      Left            =   5280
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtValorTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7560
      TabIndex        =   3
      Text            =   " "
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtToTDesc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9360
      TabIndex        =   2
      Text            =   " "
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5040
      TabIndex        =   1
      Top             =   1560
      Width           =   5775
   End
   Begin MSMask.MaskEdBox txtDtEmis 
      Height          =   370
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":739E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":76BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":7B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":7F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":8978
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":938A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":9D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fincancelalanc.frx":A7AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   1270
      ButtonWidth     =   2646
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
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
            Caption         =   "&Cancelar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   5
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
         Left            =   9720
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fincancelalanc.frx":B1C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fincancelalanc.frx":C35A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fincancelalanc.frx":D3E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fincancelalanc.frx":EAE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fincancelalanc.frx":FBF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fincancelalanc.frx":10BA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fincancelalanc.frx":11D74
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTAASS 
      Height          =   5025
      Left            =   60
      TabIndex        =   8
      Top             =   2160
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   8864
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Modalidade"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Venc."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Dt.Baixa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor Desc."
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Histórico"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   20
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
   Begin MSMask.MaskEdBox txtBaixa 
      Height          =   370
      Left            =   3600
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data Emissão:"
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
      Left            =   1920
      TabIndex        =   17
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nº Lançamento:"
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
      TabIndex        =   16
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Status Lançamento:"
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
      Left            =   5280
      TabIndex        =   15
      Top             =   840
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00808080&
      Height          =   1335
      Left            =   45
      Top             =   720
      Width           =   10875
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Valor Total:"
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
      Left            =   7560
      TabIndex        =   14
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total Desconto:"
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
      Left            =   9360
      TabIndex        =   13
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cliente/Fornecedor:"
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
      Left            =   330
      TabIndex        =   12
      Top             =   1650
      Width           =   1620
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Data Baixa:"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   945
   End
End
Attribute VB_Name = "frmFINCANCELALANC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   cmbAuxStatusLanc.Clear
   cmbStatusLanc.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where TIPO = 'B' "
   SQL = SQL & " and codigo = '5' "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbStatusLanc.AddItem Trim(TabDESCR!DESCRICAO)
      cmbAuxStatusLanc.AddItem TabDESCR!codigo
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF9
      Case vbKeyF10
      Case vbKeyEscape
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub listaass_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTAASS, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "print"
         If txtLanc.Text <> "" Then
            FORMULA_REL = "{ITEMLANCAMENTO.numr_doc} = " & txtLanc.Text

            If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

            Nome_Relatorio = "rel_fin01.rpt"
            frmRELATORIO10.Show 1
         End If
      Case "matar"
         Msg = "Confirma exclusão do lançamento ?"
         PERGUNTA Msg, vbYesNo + 32, "Exclui Lancamento", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then _
            MATA_LANCAMENTO
      Case "consultar"
         Indr_Consulta = True
         frmFINCONSLANC.Show 1
         If CRITERIO <> "" Then
            txtLanc.Enabled = True
            txtLanc.Text = CRITERIO
            txtLanc.SetFocus
         End If
         CRITERIO = ""
      Case "gravar"
         EFETIVA_GRAVA
         LIMPA_TUDO
         txtLanc.SetFocus
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
         txtLanc.SetFocus
   End Select
End Sub

Private Sub txtLanc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If IsNumeric(txtLanc.Text) Then _
         MOSTRA_LANCAMENTO
      txtDtEmis.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtDTEMIS_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtDTVENC_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtDtBaixa_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cmbstatusLanc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtCNPJCPF_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtcli_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtbaixa_GotFocus()
   txtBaixa.PromptInclude = False
   If txtBaixa.Text = "" Then
      txtBaixa.Mask = "##/##/####"
      Else
         txtBaixa.PromptInclude = True
         If Not IsDate(txtBaixa.Text) Then _
            txtBaixa.Mask = "##/##/####"
   End If
   txtBaixa.PromptInclude = True
End Sub

Private Sub txtbaixa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbStatusLanc.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtbaixa_LostFocus()
   txtBaixa.PromptInclude = False
   If txtBaixa.Text <> "" Then
      txtBaixa.PromptInclude = True
      If Not IsDate(txtBaixa.Text) Then
         txtBaixa.Mask = "##/##/####"
         txtBaixa.PromptInclude = False
            txtBaixa.Text = Date
         txtBaixa.PromptInclude = True
         Else: MsgBox "Todos itens desse lançamento serão baixados"
      End If
   End If
End Sub
'====================================
Private Sub LIMPA_TUDO()
   If IsNumeric(txtLanc.Text) Then
      SQL = "delete from LANCATEMP "
      SQL = SQL & " where numr_doc = " & txtLanc.Text
      CONECTA_RETAGUARDA.Execute SQL
   End If
   LISTAASS.ListItems.Clear
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtValorTotal.Text = ""
   txtToTDesc.Text = ""
   txtLanc.Enabled = True
   txtLanc.Text = ""
   txtDtEmis.PromptInclude = False
   txtDtEmis.Text = ""
   cmbStatusLanc.Text = ""
   cmbAuxStatusLanc.Text = ""
   txtBaixa.PromptInclude = False
   txtBaixa.Text = ""
   txtCli.Text = ""
   VALOR_TOTAL_N = 0
   VALOR_ITEM_N = 0
   VALOR_DIFERENCA_N = 0
   VLR_DESCT_DIF_N = 0
   INDR_GRAVA = False
End Sub

Private Sub MOSTRA_LANCAMENTO()

   SQL = "select * from LANCAMENTO "
   SQL = SQL & " where numr_doc = " & txtLanc.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabLancamento.EOF Then
      'CABEÇA
      If IsDate(TabLancamento!DT_LANC) Then
         TabCABECA!DT_LANC = TabLancamento!DT_LANC
         txtDtEmis.Text = TabLancamento!DT_LANC
      End If
      SQL = "select nome from CLIENTE "
      SQL = SQL & " where cgccpf='" & TabLancamento!prop & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         txtCli.Text = TabCliente!NOME
         Else
            SQL = "select * from vwFornecedor "
            SQL = SQL & " where cnpjcpf = '" & TabLancamento!prop & "'"
            TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabAUX.EOF Then _
               txtCli.Text = TabAUX!NOME
            TabAUX.Close
      End If
      TabCliente.Close
      
      
      SQL = "select * from LANCATEMP "
      SQL = SQL & " where NUMR_DOC = " & txtLanc.Text
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCABECA.EOF Then
         SqL2 = "INSERT INTO LANCATEMP (EMPRESA_ID, numr_doc, prop, dt_lanc, VALOR_LANC, TOTAL_DESCONTO, TIPO_LANCAMENTO) "
         SqL2 = SqL2 & " VALUES (" & EMPRESA_ID_N & "," & TabLancamento!NUMR_DOC & ",'" & TabLancamento!prop & "','" & DMA(TabLancamento!DT_LANC) & "'," & tpMOEDA(TabLancamento!Valor_Lanc) & "," & tpMOEDA(TabLancamento!Total_Desconto) & "," & INDR_RECEITA & ")"
         CONECTA_RETAGUARDA.Execute SqL2
      Else
         SqL2 = "UPDATE LANCATEMP SET numr_doc = " & TabLancamento!NUMR_DOC & ", prop = '" & TabLancamento!prop & "', dt_lanc = '" & DMA(TabLancamento!DT_LANC) & "',"
         SqL2 = SqL2 & " VALOR_LANC = " & tpMOEDA(TabLancamento!Valor_Lanc) & ", TOTAL_DESCONTO = " & tpMOEDA(TabLancamento!Total_Desconto) & ", TIPO_LANCAMENTO = " & INDR_RECEITA & " where NUMR_DOC = " & txtLanc.Text
         CONECTA_RETAGUARDA.Execute SqL2
      End If
      TabCABECA.Close
      
      VALOR_TOTAL_N = TabLancamento!Valor_Lanc
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = TabLancamento!prop
      VALOR_TOTAL_DESCONTO_N = TabLancamento!Total_Desconto
      txtToTDesc.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
      txtValorTotal.Text = Format(TabLancamento!Valor_Lanc, strFormatacao2Digitos)
      VALOR_TOTAL_N = TabLancamento!Valor_Lanc
      
      If Not IsNull(TabLancamento!TIPO_LANCAMENTO) Then
         If TabDESCR.State = 1 Then _
            TabDESCR.Close

         SQL = "select * from DESCR "
         SQL = SQL & " where TIPO = 'B' "
         SQL = SQL & "and codigo = '" & Trim(TabLancamento!TIPO_LANCAMENTO) & "'"
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            cmbStatusLanc.Text = Trim(TabDESCR!DESCRICAO)
            cmbAuxStatusLanc.Text = TabDESCR!codigo
         End If
         If TabDESCR.State = 1 Then _
            TabDESCR.Close
      End If
      'ITENS
      SQL = "select * from ITEMLANCAMENTO "
      SQL = SQL & " where numr_doc = " & txtLanc.Text
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabAUX.EOF
         SQL = "select * from LANCAITEMTEMP "
         SQL = SQL & " where numr_doc = " & TabAUX!NUMR_DOC
         SQL = SQL & " and seq = " & TabAUX!SEQ
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            SqL2 = "UPDATE LANCAITEMTEMP SET Valor_Item = " & tpMOEDA(TabAUX!Valor_Item) & ", Status = '" & TabAUX!Status & "', formapagto_id = " & TabAUX!FORMAPAGTO_ID & ","
            SqL2 = SqL2 & " DT_VENCIMENTO = '" & DMA(TabAUX!DT_VENCIMENTO) & "', dt_baixa = '" & DMA(TabAUX!DT_BAIXA) & "', PERC_DESCONTO = " & tpMOEDA(TabAUX!Valor_Desconto) & ", CODG_USU_BAIXA = " & TabAUX!CODG_USU_BAIXA & " where numr_doc = " & TabAUX!NUMR_DOC & " seq = " & TabAUX!SEQ
            CONECTA_RETAGUARDA.Execute SqL2
         Else
            SqL2 = "INSERT INTO LANCAITEMTEMP (numr_doc, seq, Valor_Item, Status, formapagto_id, DT_VENCIMENTO, dt_baixa, PERC_DESCONTO, PERC_JUROS, CODG_USU_BAIXA) "
            SqL2 = SqL2 & " VALUES (" & TabAUX!NUMR_DOC & "," & TabAUX!SEQ & "," & tpMOEDA(TabAUX!Valor_Item) & ",'" & TabAUX!Status & "'," & TabAUX!FORMAPAGTO_ID & ",'" & DMA(TabAUX!DT_VENCIMENTO) & ",'" & DMA(TabAUX!DT_BAIXA) & "'," & tpMOEDA(TabAUX!Valor_Desconto) & "," & TabAUX!CODG_USU_BAIXA & ")"
            CONECTA_RETAGUARDA.Execute SqL2
         End If
         TabTemp.Close
         TabAUX.MoveNext
      Wend
      TabAUX.Close
      SETA_GRID
   End If
   TabLancamento.Close
End Sub

Private Sub EFETIVA_GRAVA()
    'CABEÇA
    SQL = "select * from LANCAMENTO "
    SQL = SQL & " where NUMR_DOC = " & txtLanc.Text
    SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
    TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
    If Not TabAUX.EOF Then
       If cmbAuxStatusLanc.Text <> "" Then
          SQL = "UPDATE LANCAMENTO SET TIPO_LANCAMENTO = " & 9 & ", DT_CANCELA = '" & Now & "'"
          SQL = SQL & " where NUMR_DOC = " & txtLanc.Text
          SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
          CONECTA_RETAGUARDA.Execute SQL
       End If
       TabAUX.Close
       'ITENS
       SQL = "select * from ITEMLANCAMENTO "
       SQL = SQL & " where numr_doc = " & txtLanc.Text
       TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
       While Not TabTemp.EOF
          TabTemp!usu_alt = USUARIO_ID_N
          TabTemp!Dt_Alt = Date
          If IsNull(TabTemp!usu_cad) Then
             TabTemp!usu_cad = USUARIO_ID_N
             TabTemp!DT_CAD = Date
             Else
                If TabTemp!usu_cad = "" Then
                   TabTemp!usu_cad = USUARIO_ID_N
                   TabTemp!DT_CAD = Date
                   Else
                      If TabTemp!usu_cad <= 0 Then
                         TabTemp!usu_cad = USUARIO_ID_N
                         TabTemp!DT_CAD = Date
                      End If
                End If
          End If

          SQL = "UPDATE LANCAMENTO SET Status = '" & "B" & "', usu_alt = " & USUARIO_ID_N & ", dt_cad = '" & Now & "', CODG_USU_BAIXA = " & USUARIO_ID_N
          SQL = SQL & " where numr_doc = " & txtLanc.Text
          SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
          CONECTA_RETAGUARDA.Execute SQL

          TabTemp.MoveNext
       Wend
       MsgBox "Lançamento cancelado."
       Else: MsgBox "Número de documento inexistente."
    End If
    TabAUX.Close
End Sub

Private Sub SETA_GRID()
   LISTAASS.ListItems.Clear
   SQL = "select * from LANCAITEMTEMP "
   SQL = SQL & " where numr_doc = " & txtLanc.Text
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NUMR_SEQ_N = TabTemp!SEQ

      SQL = "select * from FORMAPAGTO "
      SQL = SQL & " where formapagto_id = " & TabTemp!FORMAPAGTO_ID
      SQL = SQL & " and status = 'true' "
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      Set Item = LISTAASS.ListItems.Add(, "seq." & TabTemp!SEQ, TabTemp!SEQ)
      Item.SubItems(1) = TabDESCR!DESCRICAO
      Item.SubItems(2) = Format(TabTemp!Valor_Item, strFormatacao2Digitos)
      Item.SubItems(3) = TabTemp!DT_VENCIMENTO
      If Not IsNull(TabTemp!DT_BAIXA) Then
         Item.SubItems(4) = TabTemp!DT_BAIXA
         Else: Item.SubItems(4) = ""
      End If
      Item.SubItems(5) = Format(TabTemp!Valor_Desconto, strFormatacao2Digitos)
      If Not IsNull(TabTemp!Status) Then
         If TabTemp!Status = "A" Then _
            Item.SubItems(6) = "Aberto"
         If TabTemp!Status = "B" Then _
            Item.SubItems(6) = "Baixado"
      End If
      SQL = "select * from OBS "
      SQL = SQL & " where prop = " & TabTemp!NUMR_DOC
      SQL = SQL & " and seq = " & TabTemp!SEQ
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then _
         If Not IsNull(TabAUX!OBS) Then _
            Item.SubItems(7) = TabAUX!OBS
      TabAUX.Close
      TabTemp.MoveNext
   Wend
   TabTemp.Close
End Sub

Public Sub MOSTRA_ITEM()
   If txtSeq.Text <> "" Then
      SQL = "select * from ITEMLANCAMENTO "
      SQL = SQL & " where numr_doc = " & txtLanc.Text
      SQL = SQL & " and seq = " & txtSeq.Text
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then
         txtValorItem.Text = Format(TabAUX!Valor_Item, strFormatacao2Digitos)
         If Not IsNull(TabAUX!Status) Then
            If TabAUX!Status = "A" Then _
               cmbStatusLancItem.Text = "Aberto"
            If TabAUX!Status = "B" Then _
               cmbStatusLancItem = "Baixado"
         End If
         SQL = "select * from FORMAPAGTO "
         SQL = SQL & " where formapagto_id = " & TabAUX!FORMAPAGTO_ID
         SQL = SQL & " and status = 'true' "
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then
            cmbAuxModalidade.Text = TabAUX!FORMAPAGTO_ID
            cmbModalidade.Text = TabDESCR!DESCRICAO
         End If
         TabDESCR.Close
         txtDtVenc.PromptInclude = False
            txtDtVenc.Text = TabAUX!DT_VENCIMENTO
         txtDtVenc.PromptInclude = True
         txtDtBaixa.PromptInclude = False
         If IsDate(TabAUX!DT_BAIXA) Then _
            txtDtBaixa.Text = TabAUX!DT_BAIXA
         txtDtBaixa.PromptInclude = True
         SQL = "select * from OBS "
         SQL = SQL & " where prop = " & TabAUX!NUMR_DOC
         SQL = SQL & " and seq = " & TabAUX!SEQ
         TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabDESCR.EOF Then _
            If Not IsNull(TabDESCR!OBS) Then _
               txtHist.Text = TabDESCR!OBS
         TabDESCR.Close

         VALOR_DESCONTO_N = TabAUX!Valor_Desconto
         'VALOR_TOTAL_DESCONTO_N = TABAUX!valor_Desconto
         VALOR_DIFERENCA_N = TabAUX!Valor_Item
         VLR_DESCT_DIF_N = VALOR_DESCONTO_N

         txtDesconto.Text = Format(VALOR_DESCONTO_N, strFormatacao2Digitos)
      End If
      TabAUX.Close
   End If
End Sub

Private Sub MATA_LANCAMENTO()
   If txtLanc.Text <> "" Then
      If IsNumeric(txtLanc.Text) Then
         SQL = "select * from LANCAMENTO "
         SQL = SQL & " where numr_doc = " & txtLanc.Text
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabLancamento.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabLancamento.EOF Then
            SQL = "select * from CABECANOTA "
            SQL = SQL & " where numr_doc = " & TabLancamento!NUMR_DOC
            TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabNOTA.EOF Then
               MsgBox "Impossível excluir lançamento, existe nota fiscal emitida, número = " & TabNOTA!NUMR_NOTA
               TabNOTA.Close
               Exit Sub
            End If
            TabNOTA.Close
            TabLancamento.Delete
         End If
         TabLancamento.Close
         LIMPA_TUDO
         txtLanc.SetFocus
      End If
   End If
End Sub
