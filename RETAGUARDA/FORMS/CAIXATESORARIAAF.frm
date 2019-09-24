VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCAIXATESORARIAAF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Abertura Caixa Tesoraria"
   ClientHeight    =   1560
   ClientLeft      =   4080
   ClientTop       =   3015
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CAIXATESORARIAAF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleMode       =   0  'User
   ScaleWidth      =   7006.151
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   -300
      TabIndex        =   2
      Top             =   700
      Width           =   8055
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtDtDig 
         Height          =   450
         Left            =   2520
         TabIndex        =   0
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   794
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Inicial R$ = "
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
         Left            =   585
         TabIndex        =   5
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Abertura : "
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
         Left            =   630
         TabIndex        =   3
         Top             =   360
         Width           =   1785
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXATESORARIAAF.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CAIXATESORARIAAF.frx":6DAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   1270
      ButtonWidth     =   2910
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "$Abrir Caixa"
            Key             =   "abrir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmCAIXATESORARIAAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim CAIXA_ID_N            As Long
   Dim VALR_SALDO_ATUAL_N    As Double
   Dim VALR_SALDO_ANTERIOR_N As Double

Private Sub Form_Load()
   'Call CentralizaJanela2(frmCAIXATESORARIAAF)

   INICIALIZA_TELA
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF10
         ABRE_CAIXA
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
   Select Case Button.key
      Case "abrir"
         MOSTRA_RODAPE "Aguarde, Processando ...", "", "", "", ""

         ABRE_CAIXA
      Case "sair"
         Unload Me
   End Select
End Sub

Private Sub txtdtdig_GotFocus()
   txtDtDig.PromptInclude = True
End Sub

Private Sub txtdtdig_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtdtdig_LostFocus()
   txtDtDig.PromptInclude = True
   If Not IsDate(txtDtDig.Text) Then
      txtDtDig.PromptInclude = False
         txtDtDig.Text = Date
      txtDtDig.PromptInclude = True
   End If
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_KeyPress"
End Sub

Private Sub FECHA_CAIXA_DIA_ANTERIOR()
'On Error GoTo ERRO_TRATA

   NUMR_ID_N = 0

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   DATA_FIM = DMA(txtDtDig.Text)
   DATA_FIM = DATA_FIM - 1

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select * from CAIXATESORARIA WITH (NOLOCK)"
   'SQL = SQL & " where dt_abertura <= '" & DATA_FIM & "'"
   SQL = SQL & " where status = 'A'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCAIXA.EOF
      NUMR_ID_N = TabCAIXA.Fields("CAIXATESORARIA_ID").Value
      DATA_INI = DMA(TabCAIXA.Fields("dt_abertura").Value)

      If IsNull(TabCAIXA!dt_fechamento) Then
         FECHA_CAIXA
         Else
            If Not IsDate(TabCAIXA!dt_fechamento) Then
               FECHA_CAIXA
               Else
                  If CDate(TabCAIXA!dt_fechamento) < CDate(1 / 1 / 2000) Then
                     FECHA_CAIXA
                  End If
            End If
      End If
      TabCAIXA.MoveNext
   Wend
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FECHA_CAIXA_DIA_ANTERIOR"
End Sub

Sub FECHA_CAIXA()
'On Error GoTo ERRO_TRATA

   Msg = "Caixa dia : " & DATA_INI & " não foi fechado."
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      SqL2 = DATA_INI
      SQL3 = DATA_FIM

      SQL = "UPDATE CAIXATESORARIA SET "
      SQL = SQL & " dt_fechamento = '" & DMA(Date) & "'"
      SQL = SQL & ", Status = '" & "F" & "'"

      'SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
      'SQL = SQL & " and dt_abertura = '" & DMA(SqL2) & "'"
      'SQL = SQL & " and CAIXATESORARIA_ID = " & NUMR_ID_N

      SQL = SQL & " where CAIXATESORARIA_ID = " & NUMR_ID_N

      CONECTA_RETAGUARDA.Execute SQL

      DATA_INI = TabCAIXA!DT_ABERTURA
      CAIXA_ID_N = TabCAIXA!CAIXATESORARIA_ID
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FECHA_CAIXA"
End Sub

Sub INICIALIZA_TELA()
'On Error GoTo ERRO_TRATA

   FECHA_CAIXA_DIA_ANTERIOR

   Dim DATA_BUSCA_DIA As String

   CAIXA_ID_N = 0
   VALR_SALDO_ATUAL_N = 0
   VALR_SALDO_ANTERIOR_N = 0

   If USUARIO_ID_N <= 0 Then
      MsgBox "Usuário inexistente."
      Unload Me
      Exit Sub
   End If
   Me.Height = 2505

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   'saldo atual caixa
   SQL = "select * from CAIXATESORARIA WITH (NOLOCK)"
   SQL = SQL & " where dt_abertura >= '" & DMA(Date, "I") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(Date, "F") & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then
      Me.Height = 1935
      Me.Caption = "Fechamento caixa Tesouraria"
      Toolbar1.Buttons(3).Caption = "Fechar Caixa"
      Label1.Caption = "Dt Fechamento:"

      CAIXA_ID_N = TabCAIXA!CAIXATESORARIA_ID

      'CREDITOS
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select sum(VALOR) from CAIXATESORARIAITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN FORMAPAGTO"
      SQL = SQL & " ON CAIXATESORARIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_ID_N
      SQL = SQL & " and tipo = 'C' "   'creditos
      'SQL = SQL & " and formapagto_id = 1 " 'SOMENTE DINHEIRO
      SQL = SQL & " and contab_tesora = 'true' "

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            If TabTemp.Fields(0).Value > 0 Then _
               VALR_SALDO_ATUAL_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close

      'DEBITOS
      SQL = "select sum(VALOR) from CAIXATESORARIAITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " ON CAIXATESORARIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_ID_N
      SQL = SQL & " and tipo = 'D' "   'debitos
      'SQL = SQL & " and formapagto_id = 1 " 'SOMENTE DINHEIRO
      SQL = SQL & " and contab_tesora = 'true' "

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            If TabTemp.Fields(0).Value > 0 Then _
               VALR_SALDO_ATUAL_N = VALR_SALDO_ATUAL_N - TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
      Else
         Me.Caption = "Abertura caixa Tesouraria"
         Toolbar1.Buttons(3).Caption = "Abrir Caixa"
   End If
'==================
   CAIXA_ID_N = 0
'verificar transpote do valor do dia anterior
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close
'aqui pego o ultimo registro da tabela, este é o ultimo dia de movimento de caixa
   SQL = "select MAX(CAIXATESORARIA_ID) from CAIXATESORARIA WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then _
      If Not IsNull(TabCAIXA.Fields(0).Value) Then _
         If TabCAIXA.Fields(0).Value > 0 Then _
            CAIXA_ID_N = TabCAIXA.Fields(0).Value
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   'saldo anterior caixa
   SQL = "select * from CAIXATESORARIA WITH (NOLOCK)"
   SQL = SQL & " where caixatesoraria_id = " & CAIXA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then
      CAIXA_ID_N = TabCAIXA!CAIXATESORARIA_ID
      DATA_BUSCA_DIA = DMA(TabCAIXA.Fields("dt_abertura").Value)

      If TabTemp.State = 1 Then _
         TabTemp.Close
'CREDITO
      SQL = "select sum(VALOR) from CAIXATESORARIAITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " ON CAIXATESORARIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_ID_N
      SQL = SQL & " and tipo = 'C' "
      'SQL = SQL & " and formapagto_id = 1 " 'SOMENTE DINHEIRO
      SQL = SQL & " and contab_tesora = 'true' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            VALR_SALDO_ANTERIOR_N = TabTemp.Fields(0).Value
      If TabTemp.State = 1 Then _
         TabTemp.Close
'DEBITO
      SQL = "select sum(VALOR) from CAIXATESORARIAITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " ON CAIXATESORARIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_ID_N
      SQL = SQL & " and tipo = 'D' "
      'SQL = SQL & " and formapagto_id = 1 " 'SOMENTE DINHEIRO
      SQL = SQL & " and contab_tesora = 'true' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields(0).Value) Then
            If TabTemp.Fields(0).Value > 0 Then
               VALR_SALDO_ANTERIOR_N = VALR_SALDO_ANTERIOR_N - TabTemp.Fields(0).Value
               Else: VALR_SALDO_ANTERIOR_N = VALR_SALDO_ANTERIOR_N + TabTemp.Fields(0).Value
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   If VALR_SALDO_ATUAL_N = 0 Then _
      VALR_SALDO_ATUAL_N = VALR_SALDO_ANTERIOR_N

   If (VALR_SALDO_ATUAL_N - VALR_SALDO_ANTERIOR_N) < 0 Then
      frmCAIXATESORARIAAF.BackColor = vbRed
      Label1.ForeColor = vbRed
      Label1.FontBold = True
      Label1.Refresh
   End If

   txtDtDig.PromptInclude = False
      txtDtDig.Text = Date
   txtDtDig.PromptInclude = True

   VALR_SALDO_ANTERIOR_N = VALR_SALDO_ANTERIOR_N + BUSCA_FAT_DIA(DATA_BUSCA_DIA)

   txtValor.Text = Format(VALR_SALDO_ANTERIOR_N, strFormatacao2Digitos)
   txtValor.Enabled = False

   MOSTRA_RODAPE "ESC - Sair", "F10 - Gravar", "", "", ""

   'If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   'SQL = "select * from CAIXATESORARIA WITH (NOLOCK)"
   'SQL = SQL & " where dt_fechamento = '" & DMA(txtDtDig.Text) & "'"
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   'TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   'If Not TabCAIXA.EOF Then
   '   If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

   '   MsgBox "Caixa Tesoraria já foi fechado."
   '   Unload Me
   '   Exit Sub
   'End If
   'If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIALIZA_TELA"
End Sub

Sub ABRE_CAIXA()

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select * from CAIXATESORARIA WITH (NOLOCK)"
   SQL = SQL & " where dt_abertura >= '" & DMA(txtDtDig.Text, "I") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(Date, "F") & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCAIXA.EOF Then
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      CAIXA_ID_N = MAX_ID("CAIXATESORARIA_ID", "CAIXATESORARIA", "", "", "", "")

      SQL = "Insert Into CAIXATESORARIA "
      SQL = SQL & " ("
         SQL = SQL & "CAIXATESORARIA_ID, Dt_abertura, "
         SQL = SQL & " usuario_id, status,ESTABELECIMENTO_ID,numero_caixa_cpu"
      SQL = SQL & " ) "
      SQL = SQL & " Values ("
         SQL = SQL & CAIXA_ID_N                       'CAIXATESORARIA_ID
         SQL = SQL & ",'" & DMA(txtDtDig.Text) & "'"  'Dt_abertura
         SQL = SQL & "," & USUARIO_ID_N               'usuario_id
         SQL = SQL & ",'A'"                           'status
         SQL = SQL & "," & ESTABELECIMENTO_ID_N
         SQL = SQL & "," & NUMERO_CAIXA_CPU
      SQL = SQL & ")"

      CONECTA_RETAGUARDA.Execute SQL

      If Trim(txtValor.Text) <> "" Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from CAIXATESORARIAITEM WITH (NOLOCK)"
         SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_ID_N
         SQL = SQL & " and CAIXATESORARIAITEM_ID = 9999 "
         SQL = SQL & " and formapagto_id = 9999 "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "INSERT INTO CAIXATESORARIAITEM "
            SQL = SQL & " (CAIXATESORARIA_ID, Valor, CAIXATESORARIAITEM_ID, formapagto_id, historico, "
            SQL = SQL & " Tipo, numr_doc, origem, Status )"
            SQL = SQL & " VALUES ("
               SQL = SQL & CAIXA_ID_N                                   'CAIXATESORARIA_ID
               SQL = SQL & "," & tpMOEDA(txtValor.Text)                 'Valor
               SQL = SQL & "," & 9999                                   'CAIXATESORARIAITEM_ID
               SQL = SQL & "," & 9999                                   'formapagto_id
               SQL = SQL & ",'" & "Valor Inicial Caixa" & "'"           'historico
               SQL = SQL & ",'C'"                                       'Tipo
               SQL = SQL & ",'01'"                                      'numr_doc
               SQL = SQL & ",'" & "T" & "'"                             'origem
               SQL = SQL & ",'" & "A" & "'"                             'Status
            SQL = SQL & " )"

            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If
      MsgBox "Caixa Tesoraria aberto com sucesso!", vbExclamation, "MEGASIM"
      Else
         Msg = "Deseja Fechar o Caixa?"
         PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
         If RESPOSTA = vbYes Then
            If TabCAIXA.State = 1 Then _
               TabCAIXA.Close

            SQL = "select * from CAIXATESORARIA WITH (NOLOCK)"
            SQL = SQL & " where dt_abertura >= '" & DMA(txtDtDig.Text, "I") & "'"
            SQL = SQL & " and dt_abertura <= '" & DMA(Date, "F") & "'"
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCAIXA.EOF Then
               SQL = "UPDATE CAIXATESORARIA SET "
               SQL = SQL & " dt_fechamento = '" & DMA(txtDtDig.Text) & "'"
               SQL = SQL & ", Status = 'F' "

               SQL = SQL & " where caixatesoraria_id = " & TabCAIXA.Fields("caixatesoraria_id").Value
               
               CONECTA_RETAGUARDA.Execute SQL
               Else: MsgBox "Caixa não foi aberto para a Data " & txtDtDig.Text
            End If
            If TabCAIXA.State = 1 Then _
               TabCAIXA.Close
            MsgBox "Caixa fechado com sucesso!", vbExclamation, "MEGASIM"
         End If
   End If
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   Unload Me
End Sub

Function BUSCA_FAT_DIA(DIA_CONSULTA As String) As Double
'On Error GoTo ERRO_TRATA

   BUSCA_FAT_DIA = 0

   Dim TabFAT As New ADODB.Recordset

   SQL = "select SUM(ITEMLANCAMENTO.VALOR_ITEM) AS ValorDia "
   SQL = SQL & " from LANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK)"
   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID "
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON ITEMLANCAMENTO.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " Where LANCAMENTO.tipo_lancamento = 1 " 'contas a receber
   SQL = SQL & " and itemLANCAMENTO.status = 'B'"
   SQL = SQL & " and itemLANCAMENTO.dt_baixa >= '" & DMA(DIA_CONSULTA, "i") & "'"
   SQL = SQL & " and itemLANCAMENTO.dt_baixa <= '" & DMA(DIA_CONSULTA, "f") & "'"

   SQL = SQL & " AND itemLANCAMENTO.FORMAPAGTO_ID = 1 " 'somente dinheiro

   SQL = SQL & " and LANCAMENTO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and contab_tesora = 'true' "

   TabFAT.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFAT.EOF Then _
      If Not IsNull(TabFAT.Fields(0).Value) Then _
         BUSCA_FAT_DIA = TabFAT.Fields(0).Value
   If TabFAT.State = 1 Then _
      TabFAT.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "BUSCA_FAT_DIA"
End Function
