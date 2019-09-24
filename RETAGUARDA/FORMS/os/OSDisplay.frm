VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSDisplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordem Serviços Pendêntes"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OSDisplay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbSituacaoAUX 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   1080
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton optDtAbertura 
      Caption         =   "Dt.Aber.:"
      Height          =   240
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton optDtFechamento 
      Caption         =   "Dt.Fech.:"
      Height          =   240
      Left            =   2880
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbSituacao 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Selecione situação Ordem de Serviço"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      Height          =   240
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Width           =   1455
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
      Picture         =   "OSDisplay.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1020
   End
   Begin MSComctlLib.ListView lstOS 
      Height          =   4785
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   5460
      _ExtentX        =   9631
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
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "O.S."
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Placa/Eqp."
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   3919
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstProduto 
      Height          =   1665
      Left            =   5520
      TabIndex        =   5
      Tag             =   "Produtos O.S."
      Top             =   720
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   2937
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qtde."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DtGarantia"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Marca"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView lstServico 
      Height          =   1665
      Left            =   5520
      TabIndex        =   6
      Tag             =   "Serviços O.S."
      Top             =   2400
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   2937
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Serviço"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Responsável"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstTermo 
      Height          =   1065
      Left            =   5520
      TabIndex        =   7
      Tag             =   "Observações O.S."
      Top             =   5520
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   1879
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Termo Garantia"
         Object.Width           =   19599
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   6000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   12105
      DesignHeight    =   6645
   End
   Begin Threed.SSCommand cmdOS 
      Height          =   900
      Left            =   1200
      TabIndex        =   8
      Top             =   5640
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1588
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "OSDisplay.frx":62E2
      Caption         =   "&O.S."
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin MSComctlLib.ListView lstOBs 
      Height          =   1425
      Left            =   5520
      TabIndex        =   9
      Tag             =   "Serviços O.S."
      Top             =   4080
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   2514
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "OBS"
         Object.Width           =   195987
      EndProperty
   End
   Begin Threed.SSCommand cmdOBS 
      Height          =   900
      Left            =   2280
      TabIndex        =   10
      Top             =   5640
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1588
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "OSDisplay.frx":6734
      Caption         =   "O&bs"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin Threed.SSCommand cmdConsultar 
      Height          =   900
      Left            =   3360
      TabIndex        =   18
      Top             =   5640
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1588
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "OSDisplay.frx":6A4E
      Caption         =   "Consulta"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdClonar 
      Height          =   900
      Left            =   4440
      TabIndex        =   19
      Top             =   5640
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1588
      _Version        =   262144
      CaptionStyle    =   1
      PictureFrames   =   1
      Picture         =   "OSDisplay.frx":8758
      Caption         =   "Clonar"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Dt.Inicial:"
      Height          =   240
      Left            =   6240
      TabIndex        =   16
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Dt.Final:"
      Height          =   240
      Left            =   9240
      TabIndex        =   15
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situação:"
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   0
      X2              =   12120
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmOSDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   cmbSituacao.Clear

   SQL = "select * from DESCR "
   SQL = SQL & " where tipo = 'Z' "
   SQL = SQL & "order by codigo "
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF
      cmbSituacao.AddItem Trim(TabDESCR!DESCRICAO)
      cmbSituacaoAUX.AddItem Trim(TabDESCR.Fields("CODIGO").Value)
      TabDESCR.MoveNext
   Wend
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   MOSTRA_OS

End Sub

Private Sub cmdConsultar_Click()
   MOSTRA_OS
End Sub

Private Sub cmbSituacao_GotFocus()
   cmbSituacao.SelStart = 0
   cmbSituacao.SelLength = Len(cmbSituacao.Text)
   cmbSituacao.BackColor = &HC0FFFF
End Sub

Private Sub cmbSituacao_Click()
On Error Resume Next

   cmbSituacaoAUX.ListIndex = cmbSituacao.ListIndex
   If Trim(cmbSituacaoAUX.Text) <> "" Then
      optDtAbertura.Value = True
      txtDtIni.PromptInclude = False
         txtDtIni.Text = DMA(Date, "i")
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = False
         txtDtFim.Text = DMA(Date, "f")
      txtDtFim.PromptInclude = True
   End If

Err.Clear
End Sub

Private Sub cmbSituacao_LostFocus()
   cmbSituacao.BackColor = &HFFFFFF
   'If Trim(cmbSITUACAO.Text) = "" Then _
      cmbSITUACAO.ListIndex = 0
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

   txtDtIni.SelStart = 0
   txtDtIni.SelLength = Len(txtDtIni.Text)
   txtDtIni.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
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
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtIni_LostFocus()
   txtDtIni.BackColor = &HFFFFFF
End Sub

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

   txtDtFim.SelStart = 0
   txtDtFim.SelLength = Len(txtDtFim.Text)
   txtDtFim.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_GotFocus"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtDtFim_LostFocus()
   txtDtFim.BackColor = &HFFFFFF
End Sub

Private Sub cmdSAIRDESCR_Click()
   Unload Me
End Sub

Private Sub cmdOS_Click()
'On Error GoTo ERRO_TRATA

   INDR_RECEITA = 1

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text
         frmOSServico.txtOs.Text = OS_ID_N
      End If
   End If

   frmOSServico.Show 1
   INDR_RECEITA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdOS_Click"
End Sub

Private Sub cmdOBS_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text

         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select * from OS "
         SQL = SQL & " where os_id = " & OS_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            CHAMADA_A = "OBS"

            frmOBS.Show 1
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdOS_Click"
End Sub

Private Sub LSTOS_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text
         SQL = "" & lstOS.SelectedItem.ListSubItems.item(3).Text
         IMPRIMIR_ORDEM_SERVICO OS_ID_N, "SERVIÇO", Trim(SQL)
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstOS_DblClick"
End Sub

Private Sub LSTTERMO_DblClick()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text

         If TabCabeca.State = 1 Then _
            TabCabeca.Close

         SQL = "select * from OS "
         SQL = SQL & " where os_id = " & OS_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabeca.EOF Then
            CHAMADA_A = "TERMOGARANTIA"

            frmOBS.Show 1

            MOSTRA_OBSTERMO
            Else
               If TabCabeca.State = 1 Then _
                  TabCabeca.Close
               MsgBox "O.S. não informada !!!"
         End If
         If TabCabeca.State = 1 Then _
            TabCabeca.Close
      End If
      CHAMADA_A = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LSTTERMO_DblClick"
End Sub

Private Sub LSTos_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text
         lstProduto.ListItems.Clear
         lstServico.ListItems.Clear
         lstOBs.ListItems.Clear
         lstTermo.ListItems.Clear
         MOSTRA_PRODUTO
         MOSTRA_SERVICO
         MOSTRA_OSOBS
         MOSTRA_OBSTERMO
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LSTos_Click"
End Sub

Sub MOSTRA_OS()
'On Error GoTo ERRO_TRATA

   Dim SIT_OS_N         As Long
   Dim SITUACAO_OS_A    As String

   SITUACAO_OS_A = ""
   OS_ID_N = 0

SITUACAO_OS_A = "1,3"

   lstOS.Visible = False
   lstOS.ListItems.Clear

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select os_id,nome_eqp,placa,cliente,dt_os,situacao_os,DESCRICAOPESSOA"
   SQL = SQL & " from vwOS WITH (NOLOCK) "

   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N

   If Trim(cmbSituacao.Text) = "" Then
      SQL = SQL & " and situacao_os in (" & SITUACAO_OS_A & ")"
      Else
         If Trim(cmbSituacaoAUX.Text) <> "" Then _
            If IsNumeric(cmbSituacaoAUX.Text) Then _
               SQL = SQL & " and situacao_os = " & cmbSituacaoAUX.Text
   End If

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      If optDtAbertura.Value = True Then
         SQL = SQL & " and dt_os >= '" & txtDtIni.Text & "'"
         SQL = SQL & " and dt_os <= '" & txtDtFim.Text & "'"
      End If
      If optDtFechamento.Value = True Then
         SQL = SQL & " and dt_fecha >= '" & txtDtIni.Text & "'"
         SQL = SQL & " and dt_fecha <= '" & txtDtFim.Text & "'"
      End If
   End If

   SQL = SQL & " order by dt_os DESC "

'Debug.Print SQL

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      If OS_ID_N <> TabCabeca.Fields("os_id").Value Then
         OS_ID_N = TabCabeca.Fields("os_id").Value
      SIT_OS_N = 0 & TabCabeca.Fields("situacao_os").Value

      Set item = lstOS.ListItems.Add(, "seq." & TabCabeca.Fields("os_id").Value, TabCabeca.Fields("os_id").Value)

      item.SubItems(1) = "" & Trim(TabCabeca.Fields("placa").Value)
      If Trim(TabCabeca.Fields("placa").Value) = "" Then _
         item.SubItems(1) = "" & Trim(TabCabeca.Fields("nome_eqp").Value)

      item.SubItems(2) = "" & Trim(TabCabeca.Fields("nome_eqp").Value)

      If Left(Trim(TabCabeca.Fields("cliente").Value), 10) = "CONSUMIDOR" Then
         item.SubItems(3) = "" & Trim(TabCabeca.Fields("DESCRICAOPESSOA").Value)
         Else: item.SubItems(3) = "" & Trim(TabCabeca.Fields("cliente").Value)
      End If
      
      item.SubItems(4) = "" & TRAZ_DESCRITOR("Z", Str(SIT_OS_N))

      If DMA(TabCabeca.Fields("dt_os").Value) = Date Then
         item.ForeColor = vbBlue
         item.ListSubItems(1).ForeColor = vbBlue
         item.ListSubItems(2).ForeColor = vbBlue
         item.ListSubItems(3).ForeColor = vbBlue
         item.ListSubItems(4).ForeColor = vbBlue
         Else
            item.ForeColor = vbRed
            item.ListSubItems(1).ForeColor = vbRed
            item.ListSubItems(2).ForeColor = vbRed
            item.ListSubItems(3).ForeColor = vbRed
            item.ListSubItems(4).ForeColor = vbRed
      End If
      End If
      TabCabeca.MoveNext
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstOS.Refresh
   lstOS.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OS"
End Sub

Sub MOSTRA_PRODUTO()
'On Error GoTo ERRO_TRATA

   lstProduto.Visible = False
   lstProduto.ListItems.Clear
   CONT_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select OSPECA.OSPECA_ID, OSPECA.OS_ID, OSPECA.PRODUTO_ID, OSPECA.DT_CAD, OSPECA.SOLICITANTE_ID, OSPECA.VALOR_ITEM, OSPECA.DESCONTO_PRODUTO, OSPECA.QTDE, OSPECA.DT_GARANTIA, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.UNIDADE_MEDIDA, PRODUTO.SITUACAO, PRODUTO.PRECO_CUSTO, PRODUTO.PRECO_Venda,"
   SQL = SQL & " Produto.PRECO_ATACADO , Produto.MARCA_ID"
   SQL = SQL & " FROM OSPECA  WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK) "
   SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where os_id = " & OS_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      QTDE_N = 0 & TabCabeca.Fields("qtde").Value
      VALOR_ITEM_N = 0 & TabCabeca.Fields("valor_item").Value
      Set item = lstProduto.ListItems.Add(, "seq." & CONT_N, Trim(TabCabeca.Fields("codg_produto").Value))

      item.SubItems(1) = "" & Trim(TabCabeca.Fields("descricao").Value)
      item.SubItems(2) = "" & Format(QTDE_N, strFormatacao3Digitos)
      item.SubItems(3) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      item.SubItems(4) = "" & Trim(TabCabeca.Fields("dt_garantia").Value)
      If Not IsNull(TabCabeca.Fields("marca_id").Value) Then _
         item.SubItems(5) = "" & TRAZ_DESCRITOR("W", TabCabeca.Fields("marca_id").Value)

      TabCabeca.MoveNext
      CONT_N = CONT_N + 1
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstProduto.Refresh
   lstProduto.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_PRODUTO"
End Sub

Sub MOSTRA_SERVICO()
'On Error GoTo ERRO_TRATA

   lstServico.Visible = False
   lstServico.ListItems.Clear
   CONT_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * from OSSERVICO WITH (NOLOCK) "
   SQL = SQL & " where os_id = " & OS_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      VALOR_ITEM_N = 0 & TabCabeca.Fields("valor_servico").Value
      Set item = lstServico.ListItems.Add(, "seq." & CONT_N, TabCabeca.Fields("ostarefa_id").Value)

      item.SubItems(1) = "" & Trim(TabCabeca.Fields("descricao").Value)
      item.SubItems(2) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
      If Not IsNull(TabCabeca.Fields("responsavel_id").Value) Then _
         item.SubItems(3) = "" & TRAZ_NOME_USUARIO(TabCabeca.Fields("responsavel_id").Value)

      If Not IsNull(TabCabeca.Fields("situacao").Value) Then
         If Trim(TabCabeca.Fields("situacao").Value) = "E" Then _
            item.SubItems(4) = "Execução"
         If Trim(TabCabeca.Fields("situacao").Value) = "P" Then _
            item.SubItems(4) = "Pendente"
         If Trim(TabCabeca.Fields("situacao").Value) = "O" Then _
            item.SubItems(4) = "Orçamento"
         If Trim(TabCabeca.Fields("situacao").Value) = "C" Then _
            item.SubItems(4) = "Cancelada"
         If Trim(TabCabeca.Fields("situacao").Value) = "F" Then _
            item.SubItems(4) = "Finalizada"
      End If

      TabCabeca.MoveNext
      CONT_N = CONT_N + 1
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstServico.Refresh
   lstServico.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_SERVICO"
End Sub

Sub MOSTRA_OBSTERMO()
'On Error GoTo ERRO_TRATA

   lstTermo.Visible = False
   lstTermo.ListItems.Clear
   CONT_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * FROM OSTERMO  WITH (NOLOCK) "
   SQL = SQL & " where os_id = " & OS_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      Set item = lstTermo.ListItems.Add(, "seq." & CONT_N, Trim(TabCabeca.Fields("OSTERMOOBS").Value))
      TabCabeca.MoveNext
      CONT_N = CONT_N + 1
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstTermo.Refresh
   lstTermo.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OBSTERMO"
End Sub

Sub MOSTRA_OSOBS()
'On Error GoTo ERRO_TRATA

   lstOBs.Visible = False
   lstOBs.ListItems.Clear
   CONT_N = 0

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "select * FROM OSOBS WITH (NOLOCK) "
   SQL = SQL & " where os_id = " & OS_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCabeca.EOF
      Set item = lstOBs.ListItems.Add(, "seq." & CONT_N, Trim(TabCabeca.Fields("OBS").Value))
      TabCabeca.MoveNext
      CONT_N = CONT_N + 1
   Wend
   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   lstOBs.Refresh
   lstOBs.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_OSOBS"
End Sub

Sub IMPRIMIR_ORDEM_SERVICO(NUMR_OS_N As Long, TIPO_REL As String, NOME_CLI As String)
'On Error GoTo ERRO_TRATA

   If NUMR_OS_N <= 0 Then _
      Exit Sub

   Dim TabVWOS          As New ADODB.Recordset
   Dim NOME_CT_A        As String
   Dim CGC_A            As String
   Dim RAZAO_SOCIAL_A   As String
   Dim NOME_FANT_A      As String
   Dim ENDERECO_EMP_A   As String
   Dim CEP_EMP_A        As String
   Dim COMP_EMP_A       As String
   Dim NUMERO_EMP_A     As String
   Dim BAIRRO_EMP_A     As String
   Dim CIDADE_EMP_A     As String
   Dim UF_EMP_A         As String
   Dim FONE_EMP_A       As String
   Dim FONE_CLIENTE_A   As String
   Dim FONE_RESP_A      As String
   Dim DT_FECHA_A       As String
   Dim RESPONSAVEL_A    As String
   Dim DT_GARANTIA_D    As String
   Dim COR_ID_N
   Dim MARCA_ID_N
   Dim TIPO_EQP_ID_N

   CGC_A = ""
   RAZAO_SOCIAL_A = ""
   NOME_FANT_A = ""
   NOME_CT_A = ""
   ENDERECO_EMP_A = ""
   CEP_EMP_A = ""
   COMP_EMP_A = ""
   NUMERO_EMP_A = ""
   BAIRRO_EMP_A = ""
   CIDADE_EMP_A = ""
   UF_EMP_A = ""
   FONE_EMP_A = ""
   FONE_CLIENTE_A = ""
   FONE_RESP_A = ""
   DT_FECHA_A = ""

   SQL = "delete from OSRELITEM "
   'sql=sql & " where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from OSREL "
   'sql=sql & " where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   If TabVWOS.State = 1 Then _
      TabVWOS.Close

   SQL = "select * from vwOS "
   SQL = SQL & " where os_id = " & NUMR_OS_N
   TabVWOS.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVWOS.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'CONSULTOR TECNICO
      SQL = "select nome from USUARIO "
      SQL = SQL & " where usuario_id = " & TabVWOS.Fields("ct_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            If Trim(TabConsulta.Fields(0).Value) <> "" Then _
               NOME_CT_A = "" & Trim(TabConsulta.Fields(0).Value)
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'ENDEREÇO EMPRESA
      SQL = "select ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
      SQL = SQL & " ENDERECO.NUMERO, CEP.Cidade, CEP.UF, CEP.IBGE_ID, CEP.Cep_ID"
      SQL = SQL & " from ENDERECO "
      SQL = SQL & " INNER JOIN EMPRESA "
      SQL = SQL & " ON ENDERECO.PESSOA_ID = EMPRESA.PESSOA_ID "
      SQL = SQL & " LEFT OUTER JOIN CEP "
      SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"
      SQL = SQL & " Where EMPRESA.empresa_ID = " & EMPRESA_ID_N
      SQL = SQL & " and endereco.tipo = 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         ENDERECO_EMP_A = "" & Trim(TabConsulta.Fields("rua").Value)
         CEP_EMP_A = "" & Trim(TabConsulta.Fields("cep_id").Value)
         COMP_EMP_A = "" & Trim(TabConsulta.Fields("COMPLEMENTO").Value)
         NUMERO_EMP_A = "" & Trim(TabConsulta.Fields("NUMERO").Value)
         BAIRRO_EMP_A = "" & Trim(TabConsulta.Fields("BAIRRO").Value)
         CIDADE_EMP_A = "" & Trim(TabConsulta.Fields("cidade").Value)
         UF_EMP_A = "" & Trim(TabConsulta.Fields("uf").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE EMPRESA
      SQL = "select FONE.NUMERO, FONE.DDD, FONE.LOCAL"
      SQL = SQL & " from EMPRESA "
      SQL = SQL & " INNER JOIN PESSOA "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_EMP_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_EMP_A = FONE_EMP_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_EMP_A = FONE_EMP_A & "  "

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE cliente
      SQL = "select FONE.NUMERO, FONE.DDD, FONE.LOCAL "
      SQL = SQL & " from CLIENTE "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON CLIENTE.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " Where CLIENTE.PESSOA_ID = " & TabVWOS.Fields("PESSOA_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_CLIENTE_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & "  "

         FONE_RESP_A = "" & Trim(TabConsulta.Fields("LOCAL").Value)

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'EMPRESA
      SQL = "SELECT EMPRESA_ID, PESSOA.PESSOA_ID, CNPJCPF, DESCRICAO, RAZAO"
      SQL = SQL & " FROM EMPRESA "
      SQL = SQL & " INNER JOIN PESSOA "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID"
      SQL = SQL & " Where empresa_ID = " & EMPRESA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         CGC_A = "" & Trim(TabConsulta.Fields("CNPJCPF").Value)
         RAZAO_SOCIAL_A = "" & Trim(TabConsulta.Fields("RAZAO").Value)
         NOME_FANT_A = "" & Trim(TabConsulta.Fields("descricao").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      DT_FECHA_A = "" & TabVWOS.Fields("DT_FECHA").Value
      If Trim(DT_FECHA_A) = "" Then
         DT_FECHA_A = ""
         Else: DT_FECHA_A = DMA(DT_FECHA_A)
      End If

COR_ID_N = 0 & TabVWOS.Fields("COR_ID").Value
If COR_ID_N <= 0 Then _
   COR_ID_N = 0 & TabVWOS.Fields("CorVeiculo").Value

MARCA_ID_N = 0 & TabVWOS.Fields("MARCA_ID").Value
If MARCA_ID_N <= 0 Then _
   MARCA_ID_N = 0 & TabVWOS.Fields("MarcaVeiculo").Value

TIPO_EQP_ID_N = 0 & TabVWOS.Fields("TIPO_EQP").Value
If TIPO_EQP_ID_N <= 0 Then _
   TIPO_EQP_ID_N = 0 & TabVWOS.Fields("TIPO_VEICULO_ID").Value

If Trim(NOME_CLI) = "" Then _
   NOME_CLI = "" & TRAZ_NOME_PESSOA(0, TabVWOS.Fields("CNPJCPF").Value)

      SQL = "insert into OSREL "
         SQL = SQL & "("
         SQL = SQL & "OS_ID,DT_OS,TIPO_OS,SITUACAO_OS,CONSULTOR_OS,"
         SQL = SQL & "KM_OS,PLACA_OS,estabelecimento_ID,DT_OS_FEHCA,NUMR_FROTA_OS,"
         SQL = SQL & "NOME_EMP,CNPJ_EMP,ENDERECO_EMP,NUMERO_EMP,COMPLEM_EMP,"
         SQL = SQL & "CEP_EMP,BAIRRO_EMP,CIDADE_EMP,UF_EMP,FONE_EMP,NOME_CLI,"
         SQL = SQL & "CNPJCPF_CLI,FONE_CLI,DESC_VEICULO,COR_VEICULO,MARCA_VEICULO,"
         SQL = SQL & "TIPO_VEICULO,ANO_VEICULO,MODELO_VEICULO,COMB_VEICULO,"
         SQL = SQL & "CHASSI_VEICULO,MOTOR_VEICULO,PESSOA_ID_CLIENTE,FONE_RESP"
         SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & NUMR_OS_N                                                               'OS_ID
         SQL = SQL & ",'" & DMA(TabVWOS.Fields("dt_os").Value) & "'"                         'DT_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("H", TabVWOS.Fields("tipo_os").Value) & "'"       'TIPO_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("Z", TabVWOS.Fields("SITUACAO_OS").Value) & "'"   'SITUACAO_OS
         SQL = SQL & ",'" & Trim(Left(NOME_CT_A, 20)) & "'"                                  'CONSULTOR_OS
         SQL = SQL & "," & Trim(TabVWOS.Fields("km").Value)                                  'KM_OS

         SQL = SQL & ",'" & Trim(TabVWOS.Fields("placa").Value) & "'"                        'PLACA_OS

         SQL = SQL & "," & ESTABELECIMENTO_ID_N                                              'estabelecimento_ID
         SQL = SQL & ",'" & DT_FECHA_A & "'"                                                 'DT_OS_FEHCA
         SQL = SQL & ",0"                                                                    'NUMR_FROTA_OS

         SQL = SQL & ",'" & Trim(Left(NOME_FANT_A, 100)) & "'"                               'NOME_EMP
         SQL = SQL & ",'" & Trim(CGC_A) & "'"                                                'CNPJ_EMP

         SQL = SQL & ",'" & Trim(Replace(ENDERECO_EMP_A, ",", ".")) & "'"                    'ENDERECO_EMP
         SQL = SQL & "," & Trim(NUMERO_EMP_A)                                                'NUMERO_EMP
         SQL = SQL & ",'" & Trim(Replace(COMP_EMP_A, ",", ".")) & "'"                        'COMPLEM_EMP
         SQL = SQL & ",'" & Trim(CEP_EMP_A) & "'"                                            'CEP_EMP
         SQL = SQL & ",'" & Trim(Replace(BAIRRO_EMP_A, ",", ".")) & "'"                      'BAIRRO_EMP
         SQL = SQL & ",'" & Trim(CIDADE_EMP_A) & "'"                                         'CIDADE_EMP
         SQL = SQL & ",'" & Trim(UF_EMP_A) & "'"                                             'UF_EMP
         SQL = SQL & ",'" & Trim(FONE_EMP_A) & "'"                                           'FONE_EMP

         SQL = SQL & ",'" & Trim(NOME_CLI) & "'"                                             'NOME_CLI

         SQL = SQL & ",'" & Trim(TabVWOS.Fields("CNPJCPF").Value) & "'"                      'CNPJCPF_CLI
         SQL = SQL & ",'" & Trim(FONE_CLIENTE_A) & "'"                                       'FONE_CLI
         SQL = SQL & ",'" & Trim(TabVWOS.Fields("descricaoveiculo").Value) & "'"             'DESC_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("S", Str(COR_ID_N)) & "'"                         'COR_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("W", Str(MARCA_ID_N)) & "'"                       'MARCA_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("A", Str(TIPO_EQP_ID_N)) & "'"                    'TIPO_VEICULO
         SQL = SQL & ",'" & Trim(TabVWOS.Fields("anoveiculo").Value) & "'"                   'ANO_VEICULO
         SQL = SQL & ",'" & Trim(TabVWOS.Fields("modeloveiculo").Value) & "'"                'MODELO_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("U", Str(TIPO_EQP_ID_N)) & "'"                    'COMB_VEICULO
         SQL = SQL & ",'" & Trim(TabVWOS.Fields("identificacao").Value) & "'"                'CHASSI_VEICULO
         SQL = SQL & ",'" & Trim(TabVWOS.Fields("EQUIPAMENTO_ID").Value) & "'"               'MOTOR_VEICULO
         SQL = SQL & "," & TabVWOS.Fields("PESSOA_id").Value                                 'PESSOA_ID_CLIENTE
         SQL = SQL & ",'" & Trim(FONE_RESP_A) & "'"                                          'FONE_RESP
      SQL = SQL & ")"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

'ITENS SERVIÇO
      SQL = "select * from OSSERVICO "
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "select nome from USUARIO "
         SQL = SQL & " where usuario_id = " & TabConsulta.Fields("responsavel_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)
         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSSERVICO_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'S' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            If TabItem.State = 1 Then _
               TabItem.Close

            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,"
               SQL = SQL & "DT_CAD,DESCRICAO,VALR_ITEM,VALR_DESCONTO,QTDE,"
               SQL = SQL & " RESPONSAVEL, CODG_PRODUTO "
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSSERVICO_ID").Value              'OSRELITEM_ID
               SQL = SQL & ",'S'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("responsavel_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSTAREFA_ID").Value               'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("DESCRICAO").Value) & "'"    'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_servico").Value)    'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_servico").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(1)                                            'QTDE
               SQL = SQL & ",'" & Trim(Left(RESPONSAVEL_A, 20)) & "'"                  'RESPONSAVEL
               SQL = SQL & ",''"                                                       'CODG_PRODUTO
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

'ITENS PRODUTO
      SQL = "select OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO "
      SQL = SQL & " from OSPECA "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         DT_GARANTIA_D = 0
         If Not IsNull(TabConsulta.Fields("dt_garantia").Value) Then _
            DT_GARANTIA_D = TabConsulta.Fields("dt_garantia").Value

         NOME_A = Replace(TabConsulta.Fields("DESCRICAO").Value, ",", ".")
         NOME_A = Replace(NOME_A, "'", "´")

         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "select descricao from vwVendedor "
         SQL = SQL & " where vendedor_id = " & TabConsulta.Fields("SOLICITANTE_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)
         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSPECA_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'P' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            If TabItem.State = 1 Then _
               TabItem.Close

            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,DT_CAD,DESCRICAO,"
               SQL = SQL & "VALR_ITEM,VALR_DESCONTO,QTDE,RESPONSAVEL,CODG_PRODUTO,dt_garantia"
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'OSRELITEM_ID
               SQL = SQL & ",'P'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("SOLICITANTE_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(NOME_A) & "'"                                   'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_ITEM").Value)       'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_PRODUTO").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTDE").Value)             'QTDE
               SQL = SQL & ",'" & Trim(RESPONSAVEL_A) & "'"                            'RESPONSAVEL
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("CODG_PRODUTO").Value) & "'" 'CODG_PRODUTO
               SQL = SQL & ",'" & DMA(DT_GARANTIA_D) & "'"                              'DT_garantia
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If

         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If
   If TabVWOS.State = 1 Then _
      TabVWOS.Close

'Sleep 3000

   FORMULA_REL = "{OSREL.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   FORMULA_REL = FORMULA_REL & " and {OSREL.OS_ID} = " & NUMR_OS_N

   If chkImp.Value = 1 Then
      If ESCOLHE_IMPRESSORA(NOME_BANCO_DADOS) = True Then
         If EQP_VEICULO = False Then
            Nome_Relatorio = "REL_OFICINA.rpt"
            Else: Nome_Relatorio = "REL_SERVICO.rpt"
         End If
         Nome_Relatorio = "REL_OFICINA.rpt"
         frmRELATORIO10.Show 1
      End If
      Else
         If EQP_VEICULO = False Then
            Nome_Relatorio = "REL_OFICINA.rpt"
            Else: Nome_Relatorio = "REL_SERVICO.rpt"
         End If
         Nome_Relatorio = "REL_OFICINA.rpt"
         frmRELATORIO10.Show 1
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIMIR_ORDEM_SERVICO"
End Sub

Private Sub cmdClonar_Click()
'On Error GoTo ERRO_TRATA

   If Not IsNull(lstOS.SelectedItem.Text) Then
      If IsNumeric(lstOS.SelectedItem.Text) Then
         OS_ID_N = 0 & lstOS.SelectedItem.Text

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from OS "
         SQL = SQL & " where os_id = " & OS_ID_N
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
'=======GERAR NUMERO OS_NOVA_PARA_CLONE
            Dim OS_ID_CLONE_N As Long

            GERA_PEDIDO_ID
            OS_ID_CLONE_N = 0 & PEDIDO_ID_N
'============

'============TABELA OS
            SQL = "Insert Into OS"
               SQL = SQL & " (OS_ID,ESTABELECIMENTO_ID,PESSOA_ID,CT_ID,DT_OS,DT_FECHA,TIPO_OS,SITUACAO_OS,KM,CLIENTE)"
            SQL = SQL & " Select " & OS_ID_CLONE_N & ",ESTABELECIMENTO_ID,PESSOA_ID,CT_ID,DT_OS,DT_FECHA,TIPO_OS,SITUACAO_OS,KM,CLIENTE "
            SQL = SQL & " From os"
            SQL = SQL & " Where os_id = " & OS_ID_N
            CONECTA_RETAGUARDA.Execute SQL

'============TABELA OSVEICEQP
            SQL = "Insert Into OSVEICEQP "
               SQL = SQL & " (OS_ID,VEIcULO_ID,EQUIPAMENTO_ID)"
            SQL = SQL & " Select " & OS_ID_CLONE_N & ",VEIcULO_ID,EQUIPAMENTO_ID "
            SQL = SQL & " From OSVEICEQP "
            SQL = SQL & " Where os_id = " & OS_ID_N
            CONECTA_RETAGUARDA.Execute SQL

'============TABELA OSEQUIPAMENTO
'============TABELA OSVEICULO

'============TABELA OSTERMO
            NUMR_ID_N = 0 & MAX_ID("OSTERMO_ID", "OSTERMO", "", "", "", "")
            SQL = "Insert Into OSTERMO"
               SQL = SQL & " (OS_ID,OSTERMO_ID,OSTERMOOBS)"
            SQL = SQL & " Select " & OS_ID_CLONE_N & "," & NUMR_ID_N & ",OSTERMOOBS"
            SQL = SQL & " From OSTERMO "
            SQL = SQL & " Where os_id = " & OS_ID_N
            CONECTA_RETAGUARDA.Execute SQL

'============TABELA OSOBS
            SQL = "Insert Into OSOBS"
               SQL = SQL & " (OS_ID,OSOBS_ID,DT_CAD,OBS)"
            SQL = SQL & " Select " & OS_ID_CLONE_N & ",OSOBS_ID,DT_CAD,OBS "
            SQL = SQL & " From OSOBS "
            SQL = SQL & " Where os_id = " & OS_ID_N
            CONECTA_RETAGUARDA.Execute SQL

'============TABELA OSSERVICO
            SQL = "Insert Into OSSERVICO"
               SQL = SQL & " (OS_ID,OSSERVICO_ID,OSTAREFA_ID,DT_CAD,SITUACAO,RESPONSAVEL_ID,"
               SQL = SQL & " VALOR_SERVICO,DESCRICAO,DESCONTO_SERVICO,DT_INICIO,DT_FIM,DT_FECHA)"
            SQL = SQL & " Select " & OS_ID_CLONE_N & ",OSSERVICO_ID,OSTAREFA_ID,DT_CAD,SITUACAO,RESPONSAVEL_ID,"
               SQL = SQL & " VALOR_SERVICO,DESCRICAO,DESCONTO_SERVICO,DT_INICIO,DT_FIM,DT_FECHA"
            SQL = SQL & " From OSSERVICO "
            SQL = SQL & " Where os_id = " & OS_ID_N
            CONECTA_RETAGUARDA.Execute SQL

'============TABELA OSPECA
            SQL = "Insert Into OSPECA"
               SQL = SQL & " (OS_ID,OSPECA_ID,PRODUTO_ID,DT_CAD,SOLICITANTE_ID,"
               SQL = SQL & " VALOR_ITEM,DESCONTO_PRODUTO,QTDE,DT_GARANTIA)"
            SQL = SQL & " Select " & OS_ID_CLONE_N & ",OSPECA_ID,PRODUTO_ID,DT_CAD,SOLICITANTE_ID,"
               SQL = SQL & " VALOR_ITEM,DESCONTO_PRODUTO,QTDE,DT_GARANTIA"
            SQL = SQL & " From OSPECA "
            SQL = SQL & " Where os_id = " & OS_ID_N
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         MsgBox "O.S. Clonada com sucesso !!! "
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdClonar_Click"
End Sub
