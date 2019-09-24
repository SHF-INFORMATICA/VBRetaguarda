VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCAIXAREL 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Caixa"
   ClientHeight    =   5760
   ClientLeft      =   3405
   ClientTop       =   2895
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CAIXAREL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      Height          =   240
      Left            =   3840
      TabIndex        =   13
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
      ForeColor       =   &H00400000&
      Height          =   4095
      Left            =   30
      TabIndex        =   3
      Top             =   1605
      Width           =   5775
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   5535
         Begin VB.OptionButton chkFat 
            Caption         =   "&Faturamento"
            Height          =   255
            Left            =   3840
            TabIndex        =   21
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton optSintetico 
            Caption         =   "&Sintético"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optAnalitico 
            Caption         =   "&Analítico"
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.ComboBox cmbVendAUX 
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
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbVend 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   3825
      End
      Begin Threed.SSOption optRecebimento 
         Height          =   270
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   476
         _Version        =   262144
         Caption         =   "Dt. Recebimento"
         Value           =   -1
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
         Height          =   315
         Left            =   3060
         TabIndex        =   1
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
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
      Begin Threed.SSOption optPedido 
         Height          =   270
         Left            =   3600
         TabIndex        =   12
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
         _Version        =   262144
         Caption         =   "Dt. Pedido"
      End
      Begin MSComctlLib.ListView lstForma 
         Height          =   1455
         Left            =   45
         TabIndex        =   20
         Top             =   2520
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2566
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
            Text            =   "Forma Pagto."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   176
         EndProperty
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   5760
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável:"
         Height          =   240
         Left            =   315
         TabIndex        =   15
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   1230
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
         Height          =   240
         Left            =   3060
         TabIndex        =   6
         Top             =   1230
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções"
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   30
      TabIndex        =   2
      Top             =   720
      Width           =   5775
      Begin Threed.SSOption optCaixaBalcao 
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   476
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Caixa Balcão"
         Value           =   -1
      End
      Begin Threed.SSOption optCaixaTesoraria 
         Height          =   270
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   476
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Caixa Tesoraria"
      End
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   10
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1164
      _Version        =   262144
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAIXAREL.frx":5C12
      Caption         =   "&Voltar"
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdImprimir 
      Height          =   660
      Left            =   1680
      TabIndex        =   9
      Top             =   10
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1164
      _Version        =   262144
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAIXAREL.frx":6DAC
      Caption         =   "Im&primir"
      PictureAlignment=   9
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmCAIXAREL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True
   txtDtIni.Text = DMA(Date, "I")
   txtDtFim.Text = DMA(Date, "F")

   MOSTRA_VENDEDORES
   QUALIFICA_VENDEDOR
   CARREGA_FORMAPAGTO
   chkFat.Visible = False

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Or TIPO_USUARIO = 7 Or TIPO_USUARIO = 3 Then
      cmbVend.Enabled = True
      cmbVend.Text = ""
   End If

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
      chkFat.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "form_load"
End Sub

Private Sub cmdSair_Click()
'On Error GoTo ERRO_TRATA

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdSair_Click"
End Sub

Private Sub cmbvend_Click()
'On Error GoTo ERRO_TRATA

   cmbVendAux.ListIndex = cmbVend.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbVend_Click"
End Sub

Private Sub cmbvend_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
      'Else: KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbvend_KeyPress"
End Sub

Private Sub optCaixaBalcao_Click(Value As Integer)
   Frame3.Visible = True
   optRecebimento.Value = True
   optRecebimento.Enabled = True
   optPedido.Enabled = True
   txtDtIni.SetFocus
   DoEvents
End Sub

Private Sub optCaixaTesoraria_Click(Value As Integer)
   Frame3.Visible = False
   optRecebimento.Value = False
   optPedido.Value = False
   optRecebimento.Enabled = False
   optPedido.Enabled = False
   txtDtIni.SetFocus
   DoEvents
End Sub

Private Sub optPedido_Click(Value As Integer)
   txtDtIni.SetFocus
End Sub

Private Sub optRecebimento_Click(Value As Integer)
   txtDtIni.SetFocus
End Sub

Private Sub TXTDTFIM_GotFocus()
   txtDtFim.PromptInclude = True
End Sub

Private Sub txtDtFim_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
         txtDtFim.Text = DMA(Date, "F")
      txtDtFim.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtfim_LostFocus"
End Sub

Private Sub TXTDTINI_GotFocus()
   txtDtIni.PromptInclude = True
End Sub

Private Sub txtDtIni_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtIni.PromptInclude = True
   If Not IsDate(txtDtIni.Text) Then
      txtDtIni.PromptInclude = False
         txtDtIni.Text = DMA(Date, "I")
      txtDtIni.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtIni_LostFocus"
End Sub

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
   
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub cmdImprimir_Click()
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

   If chkFat.Value = True Then
      DATA_INI = txtDtIni.Text
      DATA_FIM = txtDtFim.Text

      FORMULA_REL = "{vwRelCaixa.estabelecimento_ID} = " & ESTABELECIMENTO_ID_N
      FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.Status} <> 'C' "

      FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.tipo_lancamento} = 1 "
      FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.DT_req} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
      FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.DT_req} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"

      If Trim(SQL3) <> "" Then _
         FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.formapagto_id} in [" & SQL3 & "]"

      If Trim(cmbVend.Text) <> "" Then _
         If IsNumeric(cmbVendAux.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.vendedor_id} = " & Trim(cmbVendAux.Text)

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "rel_fat.rpt"
      frmRELATORIO10.Show 1
      Else
         If optCaixaBalcao.Value = True Then
            DATA_INI = txtDtIni.Text
            DATA_FIM = txtDtFim.Text
      
            FORMULA_REL = "{vwRelCaixa.estabelecimento_ID} = " & ESTABELECIMENTO_ID_N
            FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.Status} <> 'C' "
      
            If optRecebimento.Value = True Then
               FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.tipo_lancamento} = 1 "
      
               FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.DT_baixa} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
               FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.DT_baixa} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"
            End If
            If optPedido.Value = True Then
               FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.DT_req} >= " & "DateTime(" & Year(DATA_INI) & "," & Month(DATA_INI) & "," & Day(DATA_INI) & "," & Hour(DATA_INI) & "," & Minute(DATA_INI) & "," & Second(DATA_INI) & ")"
               FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.DT_req} <= " & "DateTime(" & Year(DATA_FIM) & "," & Month(DATA_FIM) & "," & Day(DATA_FIM) & "," & Hour(DATA_FIM) & "," & Minute(DATA_FIM) & "," & Second(DATA_FIM) & ")"
            End If
      
            If Trim(SQL3) <> "" Then _
               FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.formapagto_id} in [" & SQL3 & "]"
      
            If Trim(cmbVend.Text) <> "" Then _
               If IsNumeric(cmbVendAux.Text) Then _
                  FORMULA_REL = FORMULA_REL & " and {vwRelCaixa.vendedor_id} = " & Trim(cmbVendAux.Text)
      
            If chkImp.Value = 1 Then _
               ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
      
            If optSintetico.Value = True Then
               Nome_Relatorio = "rel_caixa_sintetico.rpt"
               Else: Nome_Relatorio = "rel_caixa_analitico.rpt"
            End If
            frmRELATORIO10.Show 1
            Else
               FORMULA_REL = "{CAIXATESORARIA.ESTABELECIMENTO_ID} = " & ESTABELECIMENTO_ID_N
               FORMULA_REL = FORMULA_REL & " and {CAIXATESORARIA.DT_Abertura} >= date (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") "
               FORMULA_REL = FORMULA_REL & " and {CAIXATESORARIA.DT_Abertura} <= date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
               FORMULA_REL = FORMULA_REL & " and {CAIXATESORARIA.Status} <> 'C' "
      
               If chkImp.Value = 1 Then _
                  ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
      
               Nome_Relatorio = "rel_caixa_tesoraria.rpt"
               frmRELATORIO10.Show 1
         End If
      End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdImprimir_Click"
End Sub

Sub QUALIFICA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select logon from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status = 1"
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      cmbVend.Enabled = False

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      CRITERIO_A = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
      SQL = "select descricao, vendedor_id from vwVendedor WITH (NOLOCK)"
      SQL = SQL & " where descricao like " & CRITERIO_A
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         cmbVend.Text = TabVENDEDOR!DESCRICAO
         cmbVendAux.Text = TabVENDEDOR!VENDEDOR_ID
      End If
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "QUALIFICA_VENDEDOR"
End Sub

Private Sub MOSTRA_VENDEDORES()
'On Error GoTo ERRO_TRADTA

   cmbVend.Enabled = True
   cmbVendAux.Enabled = True
   cmbVend.Clear
   cmbVendAux.Clear

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close
   SQL = "select descricao,vendedor_id from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabVENDEDOR.EOF
      cmbVend.AddItem Trim(TabVENDEDOR!DESCRICAO) & "-" & Trim(TabVENDEDOR!VENDEDOR_ID)
      cmbVendAux.AddItem Trim(TabVENDEDOR!VENDEDOR_ID)
      TabVENDEDOR.MoveNext
   Wend
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select logon from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
'SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and status = 1"
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      cmbVend.Enabled = False

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      CRITERIO_A = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
      SQL = "select descricao, vendedor_id from vwVendedor WITH (NOLOCK)"
      SQL = SQL & " where descricao like " & CRITERIO_A
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         cmbVend.Text = Trim(TabVENDEDOR!DESCRICAO)
         cmbVendAux.Text = Trim(TabVENDEDOR!VENDEDOR_ID)
      End If
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VENDEDORES"
End Sub

Sub CARREGA_FORMAPAGTO()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select formapagto_id,descricao from FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " where status = 'true' "
   SQL = SQL & " and contab_balcao = 'true' "
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
