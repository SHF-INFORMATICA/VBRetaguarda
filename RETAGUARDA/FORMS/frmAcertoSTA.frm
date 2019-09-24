VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAcertoSTA 
   Caption         =   "Acerto STA"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcertoSTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   12300
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   10440
      TabIndex        =   12
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtDebito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   10440
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtCredito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   10440
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
   Begin MSComctlLib.ListView lstCredito 
      Height          =   1815
      Left            =   45
      TabIndex        =   2
      Top             =   1440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3201
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Origem"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Histórico"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
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
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10080
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
               Picture         =   "frmAcertoSTA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAcertoSTA.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAcertoSTA.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAcertoSTA.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAcertoSTA.frx":9EFB
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   9840
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   840
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   840
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView lstDebito 
      Height          =   4335
      Left            =   45
      TabIndex        =   7
      Top             =   3480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7646
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   192
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Origem"
         Object.Width           =   15875
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Histórico"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   3528
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   12300
      DesignHeight    =   7905
   End
   Begin MSComctlLib.ListView lstForma 
      Height          =   1815
      Left            =   9000
      TabIndex        =   14
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3201
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Forma Pagto."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   176
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   6120
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total = "
      Height          =   240
      Index           =   2
      Left            =   9735
      TabIndex        =   13
      Top             =   4800
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   8880
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debitos = "
      Height          =   240
      Index           =   0
      Left            =   9510
      TabIndex        =   11
      Top             =   4320
      Width           =   945
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Creditos = "
      Height          =   240
      Index           =   1
      Left            =   9435
      TabIndex        =   9
      Top             =   3840
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   13200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   6120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Inicial:"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Final:"
      Height          =   240
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   915
   End
End
Attribute VB_Name = "frmAcertoSTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   CARREGA_FORMAPAGTO
   Call TXTDTINI_GotFocus
   Call TXTDTFIM_GotFocus

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         LIMPA_TUDO
      Case "consultar"
         SETA_GRID
      Case "limpar"
         LIMPA_TUDO
      Case "voltar"
         Unload Me
      Case "imprimir"
         'MONTA_REL
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstCredito_Click()
   MOSTRA_VALOR_ACERTO
End Sub

Private Sub lstDebito_Click()
   MOSTRA_VALOR_ACERTO
End Sub

Private Sub TXTDTINI_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = DMA(Date, "I")
   txtDtIni.PromptInclude = True

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

Private Sub TXTDTFIM_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = DMA(Date, "F")
   txtDtFim.PromptInclude = True

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

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim i

   SQL3 = ""
   INDR_PRI = True

   If lstForma.ListItems.Count > 0 Then
      For i = lstForma.ListItems.Count To 1 Step -1
         If lstForma.ListItems(i).Checked = True Then
            If INDR_PRI = True Then
               SQL3 = lstForma.ListItems(i).SubItems(1)
               Else: SQL3 = SQL3 & "," & lstForma.ListItems(i).SubItems(1)
            End If
            INDR_PRI = False
         End If
      Next i
   End If

   lstCredito.Visible = False
   lstCredito.ListItems.Clear
   lstDebito.Visible = False
   lstDebito.ListItems.Clear
   CONT_N = 0
   VALOR_TOTAL_N = 0
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwLeCaixaTesoraria WITH (NOLOCK) "
   SQL = SQL & " where tipo = 'C'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

SQL = SQL & " and formapagto_id in (" & Trim(SQL3) & ") "

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_abertura >= '" & Trim(txtDtIni.Text) & "'"
      SQL = SQL & " and dt_abertura <= '" & Trim(txtDtFim.Text) & "'"
   End If
   SQL = SQL & " order by dt_abertura desc"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      DoEvents
      CONT_N = CONT_N + 1

      Set item = lstCredito.ListItems.Add(, "seq." & CONT_N, TabTemp.Fields("numr_doc").Value)

      item.SubItems(1) = "" & TabTemp.Fields("historico").Value
      VALOR_ITEM_N = (TabTemp.Fields("valor").Value)
      item.SubItems(2) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

      VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N

      txtCredito.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)

      item.Checked = True

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstCredito.Visible = True

   VALOR_ITEM_N = 0
   VALOR_TOTAL_N = 0

   SQL = "select * from vwLeCaixaTesoraria WITH (NOLOCK) "
   SQL = SQL & " where tipo = 'D'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

SQL = SQL & " and formapagto_id in (" & Trim(SQL3) & ") "

   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      SQL = SQL & " and dt_abertura >= '" & Trim(txtDtIni.Text) & "'"
      SQL = SQL & " and dt_abertura <= '" & Trim(txtDtFim.Text) & "'"
   End If
   SQL = SQL & " order by dt_abertura desc"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      DoEvents
      CONT_N = CONT_N + 1

      Set item = lstDebito.ListItems.Add(, "seq." & CONT_N, TabTemp.Fields("numr_doc").Value)

      item.SubItems(1) = "" & TabTemp.Fields("historico").Value
      VALOR_ITEM_N = TabTemp.Fields("valor").Value
      item.SubItems(2) = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)

      VALOR_TOTAL_N = VALOR_TOTAL_N + VALOR_ITEM_N

      txtDebito.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)

      item.Checked = True

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

   lstDebito.Visible = True
   Me.Enabled = True
   Me.KeyPreview = True
   MOSTRA_VALOR_ACERTO

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub LIMPA_TUDO()
   txtCredito.Text = ""
   txtDebito.Text = ""
   lstCredito.ListItems.Clear
   lstDebito.ListItems.Clear
   Call TXTDTINI_GotFocus
   Call TXTDTFIM_GotFocus
End Sub

Sub MOSTRA_VALOR_ACERTO()

   Dim VALOR_CREDITO_N As Double
   Dim VALOR_DEBITO_N As Double
   Dim i             As Integer

   VALOR_CREDITO_N = 0
   VALOR_ITEM_N = 0
   INDR_PRI = False

   If lstCredito.ListItems.Count > 0 Then
      For i = lstCredito.ListItems.Count To 1 Step -1
         If lstCredito.ListItems(i).Checked = True Then
            If Trim(lstCredito.ListItems(i).Text) <> "" Then

               INDR_PRI = True
               VALOR_ITEM_N = Trim(lstCredito.ListItems(i).SubItems(2))
               VALOR_CREDITO_N = VALOR_ITEM_N + VALOR_CREDITO_N
               txtCredito.Text = Format(VALOR_CREDITO_N, strFormatacao2Digitos)
               txtCredito.Refresh

            End If
         End If '
      Next i
   End If
'====================================
   VALOR_DEBITO_N = 0
   VALOR_ITEM_N = 0
   INDR_PRI = False

   If lstDebito.ListItems.Count > 0 Then
      For i = lstDebito.ListItems.Count To 1 Step -1
         If lstDebito.ListItems(i).Checked = True Then
            If Trim(lstDebito.ListItems(i).Text) <> "" Then

               INDR_PRI = True
               VALOR_ITEM_N = Trim(lstDebito.ListItems(i).SubItems(2))
               VALOR_DEBITO_N = VALOR_ITEM_N + VALOR_DEBITO_N
               txtDebito.Text = Format(VALOR_DEBITO_N, strFormatacao2Digitos)
               txtDebito.Refresh

            End If
         End If
      Next i
   End If
'===============================================

   If Trim(txtCredito.Text) <> "" Then _
      VALOR_CREDITO_N = txtCredito.Text
   If Trim(txtDebito.Text) <> "" Then _
      VALOR_DEBITO_N = txtDebito.Text
   txtValor.Text = "" & Format(VALOR_CREDITO_N + VALOR_DEBITO_N, strFormatacao2Digitos)

End Sub

Sub CARREGA_FORMAPAGTO()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select formapagto_id,descricao from FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " where status = 'true' "
   'SQL = SQL & " and contab_balcao = 'true' "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      Set item = lstForma.ListItems.Add(, "seq." & TabConsulta.Fields(0).Value, Trim(TabConsulta.Fields("descricao").Value))
      item.SubItems(1) = "" & TabConsulta.Fields(0).Value

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   Dim i

   If lstForma.ListItems.Count > 0 Then
      For i = lstForma.ListItems.Count To 1 Step -1
         lstForma.ListItems(i).Checked = True
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_FORMAPAGTO"
End Sub
