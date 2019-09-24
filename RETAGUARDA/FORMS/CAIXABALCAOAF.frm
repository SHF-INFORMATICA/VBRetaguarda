VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCAIXABALCAOAF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura Caixa Balcão"
   ClientHeight    =   2370
   ClientLeft      =   5655
   ClientTop       =   4140
   ClientWidth     =   5295
   ForeColor       =   &H00000000&
   Icon            =   "CAIXABALCAOAF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5295
   Begin VB.CheckBox chkX 
      Caption         =   "Imprimir Leitura X"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtDolar 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtValorAbertura 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
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
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   240
         Top             =   240
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
               Picture         =   "CAIXABALCAOAF.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXABALCAOAF.frx":6DAC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtDtDig 
      Height          =   405
      Left            =   2880
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Dolar Dia US$ = "
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
      Left            =   390
      TabIndex        =   7
      Top             =   3120
      Width           =   2505
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
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
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   5370
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
      Left            =   945
      TabIndex        =   2
      Top             =   2520
      Width           =   1950
   End
   Begin VB.Label lblDATA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   1110
      TabIndex        =   1
      Top             =   1920
      Width           =   1785
   End
End
Attribute VB_Name = "frmCAIXABALCAOAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Height = 2760
   chkX.Visible = False

   MOSTRA_RODAPE "ESC - Sair", "", "", "", ""

   txtDtDig.PromptInclude = False
      txtDtDig.Text = Now 'Format(Date, "dd/mm/yyyy")
   txtDtDig.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_Activate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_Unload"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "abrir"
         MOSTRA_RODAPE "Aguarde, Processando ...", "", "", "", ""

         If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And USUARIO_ID_N = 144) Then
            SQL3 = IMPRESSORA_FISCAL_N
            CRITERIO_A = Trim(UCase(TRAZ_DESCRITOR("C", SQL3)))
            Select Case CRITERIO_A
               Case "BEMATECH"
                  ROTINA_BEMATECH
               Case "DARUMA"
                  ROTINA_DARUMA
               Case "Sweda"
                  'ROTINA_SWEDA
            End Select
         End If
'===================================
         'If Trim(txtValorAbertura.Text) <> "" Then
         '   If IsNumeric(txtValorAbertura.Text) Then
         '      INICIALIZA_TESOURARIA
         '   End If
         'End If
'============================
         MOVIMENTO_CAIXA
         frmINICIO.CHECA_CAIXA Format(Date, "dd/mm/yyyy"), "B"
         frmINICIO.BARI.Panels.Clear
         'CHECA_COMANDA_PENDENTE
      Case "sair"
         Unload Me
   End Select
End Sub

Private Sub chkX_Click()
   txtValorAbertura.SetFocus
End Sub

Private Sub txtdtdig_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDtDig.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtdtdig_GotFocus"
End Sub

Private Sub txtdtdig_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtDtDig_KeyPress"
End Sub

Private Sub txtdtdig_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtDig.PromptInclude = True
   If Not IsDate(txtDtDig.Text) Then
      txtDtDig.PromptInclude = False
         txtDtDig.Text = Format(Date, "dd/mm/yyyy")
         txtDtDig.Text = Now
      txtDtDig.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtdtdig_LostFocus"
End Sub

Private Sub txtValorAberturaAbertura_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtValorAberturaAbertura.SetFocus
      txtValorAberturaAbertura.Text = tpMOEDA(txtValorAberturaAbertura.Text)
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtValorAberturaAbertura_KeyPress"
End Sub

Private Sub txtValorAberturaAbertura_LostFocus()
'On Error GoTo ERRO_TRATA

   If txtValorAberturaAbertura.Text = "" Then _
      txtValorAberturaAbertura.Text = 0
   txtValorAberturaAbertura.Text = Format(txtValorAberturaAbertura.Text, strFormatacao2Digitos)
   txtValorAbertura.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, ""
End Sub

Sub MOVIMENTO_CAIXA()
'On Error GoTo ERRO_TRATA

   Dim Indr_Imp_Func As Boolean

   frmINICIO.CHECA_CAIXA Date, "B"  'caixa balcão
   If TabCAIXA.State = 1 Then
      If TabCAIXA.EOF Then
         If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And USUARIO_ID_N = 144) Then
            'Verifica_Cadastro_Aliquota

            Indr_Erro = False
            Indr_Imp_Func = False

            If txtValorAbertura.Text = "" Then _
               txtValorAbertura.Text = 0

            CRITERIO_A = Format(txtValorAbertura.Text, strFormatacao2Digitos)

            If chkX.Value = 1 Then _
               RETORNO_ECF = Bematech_FI_AberturaDoDia(Replace(CRITERIO_A, ",", "."), "Dinheiro")

            NUMR_ID_N = 0
            NUMR_ID_N = Bematech_FI_NumeroCaixa(NUMR_ID_N)
         End If

         '1 dólar = 1,84 reais
         'x dólares = 97 reais
         'Logo, x = 97/1,84 = 52,72 dólares

         valor_dolar_n = 0 & txtDolar.Text

         CAIXA_DIA_ID_N = MAX_ID("CAIXADIA_ID", "CAIXADIA", "", "", "", "")

         SqL2 = "INSERT INTO CAIXADIA "
         SqL2 = SqL2 & " ("
            SqL2 = SqL2 & " caixadia_id,Descricao,dt_Abertura,"
            SqL2 = SqL2 & " Tipo,Status,usuario_id,VALOR_DOLAR,estabelecimento_id,numero_caixa_cpu"
         SqL2 = SqL2 & " ) "
            SqL2 = SqL2 & " VALUES ("
            SqL2 = SqL2 & CAIXA_DIA_ID_N
            SqL2 = SqL2 & ",'" & "Caixa Balcao" & "'"
            SqL2 = SqL2 & ",'" & Now & "'"
            SqL2 = SqL2 & ",'" & "B" & "'"
            SqL2 = SqL2 & ",'" & "A" & "'"
            SqL2 = SqL2 & "," & USUARIO_ID_N
            SqL2 = SqL2 & "," & tpMOEDA(valor_dolar_n)
            SqL2 = SqL2 & "," & ESTABELECIMENTO_ID_N
            SqL2 = SqL2 & "," & NUMERO_CAIXA_CPU
         SqL2 = SqL2 & ")"
         CONECTA_RETAGUARDA.Execute SqL2

         If Not IsNumeric(txtValorAbertura.Text) Then _
            txtValorAbertura.Text = 0

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from CAIXADIA "
         SQL = SQL & " INNER JOIN CAIXADIAITEM "
         SQL = SQL & " ON CAIXADIA.CAIXADIA_ID = CAIXADIAITEM.CAIXADIA_ID "

         SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and NUMERO_CAIXA_CPU = " & NUMERO_CAIXA_CPU

         SQL = SQL & " and CAIXADIAITEM.caixadia_id = " & CAIXA_DIA_ID_N
         SQL = SQL & " and formapagto_id = 1 " 'Dinheiro
         SQL = SQL & " and usuario_id = " & USUARIO_ID_N

         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            CAIXA_DIA_ID_N = MAX_ID("CAIXADIAITEM_ID", "CAIXADIAITEM", "", "", "", "")

            SQL = "INSERT INTO CAIXADIAITEM "
               SQL = SQL & " (caixadiaitem_id,caixadia_id,Numr_Doc,VALOR_INICIAL,STATUS,TIPO_DC,HISTORICO,FORMAPAGTO_ID) "
            SQL = SQL & " VALUES ("
               SQL = SQL & CAIXA_DIA_ID_N
               SQL = SQL & "," & CAIXA_DIA_ID_N
               SQL = SQL & "," & CAIXA_DIA_ID_N
               SQL = SQL & "," & tpMOEDA(txtValorAbertura.Text)
               SQL = SQL & ",'A'"
               SQL = SQL & ",'DD'"
               SQL = SQL & ",'Abertura Caixa'"
               SQL = SQL & ","
               SQL = SQL & 1
            SQL = SQL & ")"
            CONECTA_RETAGUARDA.Execute SQL
            Else
               SQL = "UPDATE CAIXADIAITEM SET "
               SQL = SQL & " Valor_INICIAL = " & tpMOEDA(txtValorAbertura.Text)
               SQL = SQL & ", HISTORICO = '" & "Abertura Caixa" & "'"

               SQL = SQL & " and CAIXADIAITEM.caixadia_id = " & CAIXA_DIA_ID_N
               SQL = SQL & " and formapagto_id = 1 " 'Dinheiro
               SQL = SQL & " and usuario_id = " & USUARIO_ID_N
               CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         txtDtDig.Enabled = False
         txtValorAbertura.Enabled = False

         Indr_Imp_Func = True

         If USUARIO_ID_N <> 144 Then
            frmINICIO.barINI.Buttons(2).Enabled = False
            frmINICIO.mnuCAIXAbalcaoABRE.Enabled = False
         End If

         frmINICIO.mnuCAIXAbalcaoFECHA.Enabled = True
         frmINICIO.barINI.Buttons(2).Enabled = True

         Msg = "Caixa aberto com sucesso."
         Else
         'MUDOU PARA MENU, FECHAR SOBRE COMANDO DO USUÁRIO FINAL DO DIA
            'If (USA_ECF = True And INDR_CAIXA = True) Or (USA_ECF = True And usuario_id_N = 144) Then
            '   Indr_Erro = False
               'If MsgBox("Esse relatório irá executar uma redução Z. Se a impressora já estiver lacrada, não poderá mais ser usada até às 23:59:59 hs. Você realmente deseja emitir esse relatório?", vbYesNo + vbExclamation + vbDefaultButton2, "Atenção") = vbNo Then _
                  Exit Sub

               'FECHA_CAIXA_Z

            'End If

            CAIXA_DIA_ID_N = TabCAIXA.Fields("CAIXADIA_ID").Value

            SQL = "UPDATE CAIXADIA SET "
            SQL = SQL & " dt_fechamento = '" & Now & "'"
            SQL = SQL & ", Status = '" & "F" & "'"
            SQL = SQL & ", numr_reducao_z = " & Numero_Contador_Z

            SQL = SQL & " where caixadia_id = " & CAIXA_DIA_ID_N
            SQL = SQL & " and usuario_id = " & USUARIO_ID_N
            
            CONECTA_RETAGUARDA.Execute SQL

            frmINICIO.mnuCAIXAbalcaoFECHA.Enabled = False
            frmINICIO.barINI.Buttons(2).Enabled = False
            Msg = "Caixa fehado com sucesso."
      End If
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close
   End If

   MOSTRA_RODAPE Msg, "", "", "", ""

   Sleep 2000
   Unload Me

Exit Sub
ERRO_TRATA:
   MsgBox Err.description
   'TRATA_ERROS Err.Description, Me.Name, "MOVIMENTO_CAIXA"
   Err.Clear
End Sub

Public Sub Verifica_Cadastro_Aliquota()

   Dim Aliquotas As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      Aliquotas = Space(1)
      Else: Aliquotas = Space(79)
   End If

   RETORNO_ECF = Bematech_FI_RetornoAliquotas(Aliquotas)
   'Call VerificaRetornoImpressora("Alíquotas Cadastradas: ", Aliquotas, "Informações da Impressora")

   Dim Aliquota_Impressora    As Variant

   Aliquota_Impressora = Trim(Aliquotas)

   Aliquota_Impressora = Replace(Aliquota_Impressora, ",", ";")

   'ALIQUOTAS CADASTRADAS NA IMPRESSORA
   CRITERIO_A = Aliquota_Impressora

   CONT_N = 0

   Dim sLine            As String
   Dim a(0 To 50)       As String
   Dim Aliquota_Imp     As Long
   Dim Aliquota_Banco   As Long

   ParseToArray CRITERIO_A, a()

'==================Programa Aliquota
   SQL = "select distinct(aliquota_icms) from PRODUTO"
   SQL = SQL & " where aliquota_icms > 0"
   SQL = SQL & " and situacao = 'A'"
   SQL = SQL & " order by aliquota_icms"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      Aliquota_Banco = Replace(TabProduto.Fields(0).Value, "0", "")

      NUMR_SEQ_N = 0
      INDR_PRI = True

      While a(NUMR_SEQ_N) <> ""
         On Error Resume Next

         If IsNumeric(a(NUMR_SEQ_N)) <> "" Then

            Aliquota_Imp = Replace(a(NUMR_SEQ_N), "0", "")

            If Aliquota_Banco = Aliquota_Imp Then _
               INDR_PRI = False

         End If

         NUMR_SEQ_N = NUMR_SEQ_N + 1

      Wend

      If INDR_PRI = True Then
         RETORNO_ECF = Bematech_FI_ProgramaAliquota(Aliquota_Banco, 0)

         Dim INDICE_ID As Long

         INDICE_ID = MAX_ID("INDICE_ID", "indice", "", "", "", "")

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from INDICE "
         SQL = SQL & " where aliquota = " & Aliquota_Banco
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then
            SqL2 = "INSERT INTO INDICE (indice_id, impressora_id, Aliquota) "
            SqL2 = SqL2 & " VALUES (" & INDICE_ID & "," & IMPRESSORA_ID_N & "," & Aliquota_Banco & ")"
            CONECTA_RETAGUARDA.Execute SqL2
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close
End Sub

Sub ROTINA_BEMATECH()
   VERIFICA_BEMATECH_LIGADA
   If INDR_DESLIGADA = False Then
      Indr_Erro = False
      Call VerificaRetornoImpressora("Bematech_FI_AberturaDoDia", "", "Emissão de Cupom Fiscal")

      If Indr_Erro = True Then
         Indr_Erro = False
         MsgBox "Erro de comunicação com impressora."
         Exit Sub
      End If
      Indr_Erro = False
   End If
End Sub

Sub ROTINA_DARUMA()
   'Verifica se Impressora Esta Ligada
   'INTRETORNO = rVerificarImpressoraLigada_ECF_Daruma()
   If INTRETORNO <> 1 Then
      'RETORNO_ECF = Daruma_FI_LeituraX()
      Call VerificaRetornoImpressoraDaruma("", "", "Leitura X")
   End If
End Sub

Sub FECHA_CAIXA_Z()

   SQL3 = IMPRESSORA_FISCAL_N
   CRITERIO_A = Trim(UCase(TRAZ_DESCRITOR("C", SQL3)))
   Select Case CRITERIO_A
      Case "BEMATECH"
         VERIFICA_BEMATECH_LIGADA
         If INDR_DESLIGADA = False Then
            Indr_Erro = False

            RETORNO_ECF = Bematech_FI_FechamentoDoDia()
            Call VerificaRetornoImpressora("", "", "Fechamento do Dia")

            Dim Reducoes As String
            Dim TituloJanela As String

            TituloJanela = "Retorno de Informações da Impressora"

            If (LocalRetorno = "1") Then 'Grava retorno em arquivo
               Reducoes = Space(4)
               Else: Reducoes = Space(4)
            End If

            RETORNO_ECF = Bematech_FI_NumeroReducoes(Reducoes)

            Numero_Contador_Z = Reducoes
            If Numero_Contador_Z <= 0 Then
               If TabCAIXA.State = 1 Then _
                  TabCAIXA.Close
               MsgBox "Caixa não pode ser fechado, número de redução Z inválido => " & Numero_Contador_Z
               Exit Sub
            End If

            If Indr_Erro = True Then
               Indr_Erro = False
               MsgBox "Erro de comunicação com impressora."
               Exit Sub
            End If
            Indr_Erro = False
         End If
      Case "DARUMA"
         'Verifica se Impressora Esta Ligada
         'INTRETORNO = rVerificarImpressoraLigada_ECF_Daruma()
         If INTRETORNO <> 1 Then
            'RETORNO_ECF = Daruma_FI_ReducaoZ(" ", " ")
            If RETORNO_ECF <> 1 Then
               Call VerificaRetornoImpressoraDaruma("Data e Hora da Impressora: ", Now, "Informações da Impressora")
               Exit Sub
            End If
         End If
      Case "Sweda"

   End Select

End Sub

Sub INICIALIZA_TESOURARIA()
'On Error GoTo ERRO_TRATA

'COMO VAI SER MAIS DE UMA VEZ QUE VAI RODAR AQUI,
'ENTÃO CHECA PRIMEIRO SE JÁ REALIZOU A ABERTURA DA TESORARIA
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select * from CAIXATESORARIA WITH (NOLOCK)"
   SQL = SQL & " where dt_abertura >= '" & DMA(Date, "I") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(Date, "F") & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCAIXA.EOF Then
      'PEGANDO O ULTIMO REGISTRO DA TABELA QUE CARACTERIZA O ULTIMO MOVIMENTO DO CAIXA TESORARIA
      CAIXA_ID_N = 0 & MAX_ID("CAIXATESORARIA_ID", "CAIXATESORARIA", "", "", "", "") - 1
'TESTA ROTINA LINHA COM DEBAIXO
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

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

      frmCAIXATESORARIAAF.txtDtDig.PromptInclude = False
         frmCAIXATESORARIAAF.txtDtDig.Text = Date
      frmCAIXATESORARIAAF.txtDtDig.PromptInclude = True

'set ver aqui      VALR_SALDO_ANTERIOR_N = VALR_SALDO_ANTERIOR_N + frmCAIXATESORARIAAF.BUSCA_FAT_DIA(DATA_BUSCA_DIA)

      frmCAIXATESORARIAAF.txtValor.Text = Format(VALR_SALDO_ANTERIOR_N, strFormatacao2Digitos)

      CAIXA_ID_N = MAX_ID("CAIXATESORARIA_ID", "CAIXATESORARIA", "", "", "", "")

      SQL = "Insert Into CAIXATESORARIA "
      SQL = SQL & " ("
         SQL = SQL & " CAIXATESORARIA_ID, Dt_abertura, "
         SQL = SQL & " usuario_id, status,ESTABELECIMENTO_ID,"
         SQL = SQL & " numero_caixa_cpu"
      SQL = SQL & " ) "
      SQL = SQL & " Values ("
         SQL = SQL & CAIXA_ID_N                    'CAIXATESORARIA_ID
         SQL = SQL & ",'" & Now & "'"              'Dt_abertura
         SQL = SQL & "," & USUARIO_ID_N            'usuario_id
         SQL = SQL & ",'A'"                        'status
         SQL = SQL & "," & ESTABELECIMENTO_ID_N
         SQL = SQL & "," & NUMERO_CAIXA_CPU
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If   'CHECANDO SE O CAIXA JÁ FOI ABERTO, SE NÃO, ENTRA NA CONDIÇÃO ACIMA
'==================== AGORA VOU ABRIR O CAIXA COM A DATA DE HOJE
   
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CAIXATESORARIAITEM WITH (NOLOCK)"
      SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_ID_N
      SQL = SQL & " and CAIXATESORARIAITEM_ID = 9999 "
      SQL = SQL & " and formapagto_id = 9999 "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
'======================
         CAIXA_ID_N = 0
         VALR_SALDO_ATUAL_N = 0
         VALR_SALDO_ANTERIOR_N = 0
         SQL3 = ESTABELECIMENTO_ID_N

'         CAIXA_ID_N = 0 & TRAZ_ID_TABELA("CAIXATESORARIA", "CAIXATESORARIA_ID", "ESTABELECIMENTO_ID", SQL3)
      
         SQL = "select MAX(CAIXATESORARIA_ID) from CAIXATESORARIA WITH (NOLOCK)"
         SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
         TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCAIXA.EOF Then _
            If Not IsNull(TabCAIXA.Fields(0).Value) Then _
               If TabCAIXA.Fields(0).Value > 0 Then _
                  CAIXA_ID_N = TabCAIXA.Fields(0).Value
         If TabCAIXA.State = 1 Then _
            TabCAIXA.Close

         
'======================
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "INSERT INTO CAIXATESORARIAITEM "
         SQL = SQL & " (CAIXATESORARIA_ID, Valor, CAIXATESORARIAITEM_ID, formapagto_id, "
         SQL = SQL & " historico, Tipo, numr_doc, origem, Status )"
         SQL = SQL & " VALUES ("
            SQL = SQL & CAIXA_ID_N                                   'CAIXATESORARIA_ID
            SQL = SQL & "," & tpMOEDA(frmCAIXATESORARIAAF.txtValor.Text)                 'Valor
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

   If USA_ECF = True Then
      Dim NUMEROCUPOM As String
      Dim RETORNOSTATUS As String
      Dim LocalRetorno As String
      If (LocalRetorno = "1") Then 'Grava retorno em arquivo
         NUMEROCUPOM = Space(1)
         Else: NUMEROCUPOM = Space(6)
      End If

      RETORNO_ECF = Bematech_FI_NumeroCupom(NUMEROCUPOM)
   End If
'==================== set ver aqui
CONT_N = 0 & NUMEROCUPOM

   NUMR_ID_N = 1

   SQL = "select max(CAIXATESORARIAITEM_ID) from CAIXATESORARIAITEM WITH (NOLOCK) "
   SQL = SQL & " where CAIXATESORARIA_ID = " & CAIXA_DIA_ID_N
   SQL = SQL & " and CAIXATESORARIAITEM_ID < 9999 "

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         NUMR_ID_N = TabTemp.Fields(0).Value + 1
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If TabTemp.State = 1 Then _
      TabTemp.Close
'set fazer igual a abertura do caixa tesouraria
   SQL = "INSERT INTO CAIXATESORARIAITEM "
   SQL = SQL & " (CAIXATESORARIA_ID, Valor, CAIXATESORARIAITEM_ID, formapagto_id, "
   SQL = SQL & " historico, Tipo, numr_doc, origem, Status )"
   SQL = SQL & " VALUES ("
      SQL = SQL & CAIXA_ID_N                                                     'CAIXATESORARIA_ID
      SQL = SQL & "," & tpMOEDA(txtValorAbertura.Text)                           'Valor
      SQL = SQL & "," & NUMR_ID_N                                                'CAIXATESORARIAITEM_ID
      SQL = SQL & "," & 1                                                        'formapagto_id
      SQL = SQL & ",'" & "Inicio Dia " & TRAZ_NOME_USUARIO(USUARIO_ID_N) & "'"   'historico
      SQL = SQL & ",'C'"                                                         'Tipo
      SQL = SQL & ",'" & NUMEROCUPOM & "'"                                       'numr_doc
      SQL = SQL & ",'B'"                                                         'origem
      SQL = SQL & ",'" & "A" & "'"                                               'Status
   SQL = SQL & " )"
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "INICIALIZA_TESOURARIA"
End Sub

Sub CHECA_COMANDA_PENDENTE()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "SELECT PEDIDOTEMP.PEDIDO_ID, PEDIDOTEMP.ESTABELECIMENTO_ID, PEDIDOTEMP.CARTAOBARRA_ID, PEDIDOTEMP.USUARIO_ID, PEDIDOTEMP.DT_PEDIDO, "
   SQL = SQL & " CARTAOBARRA.CODIGO_BARRA, CARTAOBARRA.DESCRICAO, CARTAOBARRA.DTCAD, CARTAOBARRA.Status"
   SQL = SQL & " FROM PEDIDOTEMP "
   SQL = SQL & " LEFT OUTER JOIN CARTAOBARRA "
   SQL = SQL & " ON PEDIDOTEMP.CARTAOBARRA_ID = CARTAOBARRA.CARTAOBARRA_ID"
   SQL = SQL & " where dt_pedido <> '" & Trim(Date) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      If TabTemp.State = 1 Then _
         TabTemp.Close
      Exit Sub
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
   frmCOMANDALISTA.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "CHECA_COMANDA_PENDENTE"
End Sub
