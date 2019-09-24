VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmLOGON 
   BackColor       =   &H00400000&
   Caption         =   "SHF INFORMÁTICA"
   ClientHeight    =   3675
   ClientLeft      =   4065
   ClientTop       =   2445
   ClientWidth     =   4710
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "LOGON.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleMode       =   0  'User
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   3735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6588
      _Version        =   262144
      BackColor       =   -2147483646
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cmbEstabAux 
         BackColor       =   &H80000001&
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbEstab 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   165
         TabIndex        =   7
         ToolTipText     =   "Selecione o grupo do produto."
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox txtUsu 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   370
         Left            =   1800
         MaxLength       =   30
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   645
         Width           =   2805
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   372
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   30
         MousePointer    =   99  'Custom
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1605
         Width           =   2805
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2355
         _Version        =   262144
         BackColor       =   -2147483646
         PictureFrames   =   1
         Picture         =   "LOGON.frx":5C12
         PictureAlignment=   12
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   262144
         ForeColor       =   16777215
         BackColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Entrar"
      End
      Begin Threed.SSCommand cmdSair 
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   262144
         ForeColor       =   16777215
         BackColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Sair"
      End
      Begin VB.Label lblgrupo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Estabelecimento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   195
         TabIndex        =   8
         Top             =   2880
         Width           =   4515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmLOGON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Senha_Usuario As String
   Dim INDR_PASSA    As Boolean
   Dim TabUSU        As New ADODB.Recordset

Private Sub Form_Load()

   INDR_PASSA = False

   If INDR_CAIXA = True Then
      If (App.PrevInstance) Then
          Dim nome_tela As String
          nome_tela = App.Title
          App.Title = "Já estou em execução, frmLOGON !!!"
          'AppActivate  nome_tela
          SendKeys "%R", True
          MsgBox "Já em execução !!!"
          End
          Exit Sub
      End If
   End If

   txtUsu.Enabled = True
   txtSenha.Enabled = False
   cmdOk.Enabled = False
   INDR_FIM = True

   Me.Height = 3285
End Sub

Private Sub Form_Activate()
   If INDR_TESTE = True Then
      txtUsu.Enabled = True
      txtSenha.Enabled = True
      txtUsu.Text = "HORACIO"
      txtSenha.Text = "SHF"
      Call cmdOK_Click
   End If

   txtUsu.Enabled = True
   txtSenha.Enabled = False
   cmdOk.Enabled = False
   INDR_FIM = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If INDR_FIM = True Then
      If INDR_CAIXA = False And INDR_REMOTO = True Then _
         Call ExitWindowsEx(0, 0)

      End
      Else
         If USUARIO_ID_N = 144 Then _
            Unload Me
   End If
End Sub

Private Sub cmbestab_Click()
'On Error GoTo ERRO_TRATA

   cmbEstabAUX.ListIndex = cmbEstab.ListIndex

   If IsNumeric(cmbEstabAUX.Text) Then
      ESTABELECIMENTO_ID_N = cmbEstabAUX.Text
      If ESTABELECIMENTO_ID_N > 0 Then _
         SETA_BANCO
      Else: MsgBox "Estabelecimento incorreto."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbEstab_Click"
End Sub

Private Sub SSCommand1_Click()
   If INDR_TESTE = True Then
      txtUsu.Enabled = True
      txtSenha.Enabled = True
   End If

   txtUsu.Text = ""
   txtSenha.Text = ""
   txtUsu.SetFocus
End Sub

Private Sub txtUsu_LostFocus()
   txtUsu.Text = UCase(txtUsu.Text)
End Sub

Private Sub txtUsu_Change()
   If Trim(txtUsu.Text) <> "" Then
      If Trim(txtSenha.Text) <> "" Then _
         cmdOk.Enabled = True

      txtSenha.Enabled = True
      Else: txtSenha.Enabled = False
   End If
End Sub

Private Sub txtUsu_GotFocus()
   txtUsu.SelStart = 0
   txtUsu.SelLength = Len(txtUsu)
End Sub

Private Sub txtUsu_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Trim(txtUsu.Text) <> "" Then
            KeyAscii = 0
            txtSenha.SetFocus
        End If
    End If
End Sub

Private Sub txtSenha_GotFocus()
   txtSenha.SelStart = 0
   txtSenha.SelLength = Len(txtSenha)
End Sub

Private Sub txtSenha_Change()
   If ((Trim(txtUsu.Text) <> "") And (Trim(txtSenha.Text) <> "")) Then _
      cmdOk.Enabled = True
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ((Trim(txtUsu.Text) <> "") And (Trim(txtSenha.Text) <> "")) Then
         KeyAscii = 0
         cmdOk.Enabled = True
         Call cmdOK_Click
      End If
   End If
End Sub

Private Sub cmdSair_Click()
   End
End Sub

Public Sub cmdOK_Click()
   If ((Trim(txtUsu.Text) <> "") And (Trim(txtSenha.Text) <> "")) Then
      Senha_Usuario = Trim(txtSenha.Text)
      USU_LOGADO = Trim(txtUsu.Text)
      LOGAR_SISTEMA
   End If
End Sub

Sub LOGAR_SISTEMA()

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
   CONT_N = 0

   'ATUALIZA_ESTABELECIMENTO

   'If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTABELECIMENTOACESSO", "") = False Then
   '   SQL = "CREATE TABLE [dbo].[ESTABELECIMENTOACESSO]("
   '   SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
   '   SQL = SQL & " [USUARIO_ID] [Int] not null) ON [PRIMARY]"
   '   CONECTA_RETAGUARDA.Execute SQL

   '   SQL = "insert into ESTABELECIMENTOACESSO values(1,144)"
   '   CONECTA_RETAGUARDA.Execute SQL
   'End If
   'If EXISTE_CAMPO_TABELA("RETAGUARDA", "LiberaPercDesconto", "ESTABELECIMENTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LiberaPercDesconto bit"
'txtCasaInicioCodgProdBarra = CasaInicioCodgProdBarra
   'If EXISTE_CAMPO_TABELA("RETAGUARDA", "CasaInicioCodgProdBarra", "ESTABELECIMENTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CasaInicioCodgProdBarra int"
'DOC_FISCAL
   'If EXISTE_CAMPO_TABELA("RETAGUARDA", "DOC_FISCAL", "ESTABELECIMENTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD DOC_FISCAL bit "

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALTERA_FATURA", "ESTABELECIMENTO") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD ALTERA_FATURA bit "
      SQL = "update ESTABELECIMENTO set altera_fatura = 1 "
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "ie", "EMPRESA") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ie"

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select USUARIO.USUARIO_ID, USUARIO.SENHA, USUARIO.LOGON, USUARIO.TIPO, "
   SQL = SQL & " ESTABELECIMENTOACESSO.ESTABELECIMENTO_ID, ESTABELECIMENTO.DESCRICAO,"
   SQL = SQL & " ESTABELECIMENTO.Empresa_ID, CONTROLE_ESTOQUE, LOCALIZACAO, "
   SQL = SQL & " INDR_INDUSTRIA, LIBERA_DESCONTO,"
   SQL = SQL & " vlr_dia_compra_prod, USA_NFe, RECEBE_PEDIDO_VENDA, ESTOQUE_NEGATIVO,"
   SQL = SQL & " LEI_12741, cgc, razao_social,empresa.empresa_id,"
   SQL = SQL & " TamanhoCodgProdBarra,TamanhoPesoValorBarra,INDR_PANIFIC,peso_valor,"
   SQL = SQL & " DESCONTO_CLIENTE , DESCONTO_FUNCIONARIO, AT_VENDA_MKP,LiberaPercDesconto, "
   SQL = SQL & " CasaInicioCodgProdBarra, limpa_pedido,DOC_FISCAL,"
   SQL = SQL & " ALTERA_FATURA"

   SQL = SQL & " from USUARIO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ESTABELECIMENTOACESSO WITH (NOLOCK) "
   SQL = SQL & " ON USUARIO.USUARIO_ID = ESTABELECIMENTOACESSO.USUARIO_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK) "
   SQL = SQL & " ON ESTABELECIMENTOACESSO.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN EMPRESA WITH (NOLOCK) "
   SQL = SQL & " ON ESTABELECIMENTO.EMPRESA_ID = EMPRESA.EMPRESA_ID"

   SQL = SQL & " where logon = '" & USU_LOGADO & "'"
   SQL = SQL & " and senha = '" & Senha_Usuario & "'"
   SQL = SQL & " and status = 'true'"
   
   SQL = SQL & " order by estabelecimento.estabelecimento_id"
   
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabUSU.EOF Then
      If TabUSU.State = 1 Then _
         TabUSU.Close

      MsgBox "Usuário não cadastrado.", vbOKOnly, "ERRO !!!"
      txtSenha_GotFocus
      Else
         USU_LOGADO = "" & Trim(TabUSU.Fields(0).Value)
         USUARIO_ID_N = 0 & TabUSU.Fields("usuario_ID").Value
         TIPO_USUARIO = 0 & TabUSU.Fields("tipo").Value
         INDR_CAIXA = False

         If TIPO_USUARIO = 7 Then _
            INDR_CAIXA = True

         cmbEstab.Clear
         cmbEstabAUX.Clear

         cmbEstab.Text = "" & Trim(TabUSU.Fields("descricao").Value)
         cmbEstab.Text = "" & Trim(TabUSU.Fields("localizacao").Value)
         cmbEstabAUX.Text = "" & TabUSU.Fields("estabelecimento_id").Value
         ESTABELECIMENTO_ID_N = cmbEstabAUX.Text

         While Not TabUSU.EOF

            'cmbEstab.AddItem Trim(TabUSU.Fields("descricao").Value)
            cmbEstab.AddItem Trim(TabUSU.Fields("localizacao").Value)
            cmbEstabAUX.AddItem TabUSU.Fields("estabelecimento_id").Value
            CONT_N = CONT_N + 1

            TabUSU.MoveNext
         Wend

         If CONT_N > 1 Then
            Me.Height = 4125
            cmbEstab.SetFocus
            Else: SETA_BANCO
         End If
   End If
   If TabUSU.State = 1 Then _
      TabUSU.Close

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
End Sub

Sub SETA_BANCO()
'On Error GoTo ERRO_TRATA

   CNPJ_ESTABELECIMENTO_N = ""
   CNPJ_CRED_CARTAO_ESTAB = ""

   'If EXISTE_CAMPO_TABELA("RETAGUARDA", "VERSAO_APLICATIVO", "ESTABELECIMENTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD VERSAO_APLICATIVO nvarchar(20)"

   'txtDiasAtrazo = DiasAtrazoCliente
   'If EXISTE_CAMPO_TABELA("RETAGUARDA", "DiasAtrazoCliente", "ESTABELECIMENTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD DiasAtrazoCliente INT "

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "USA_TAB_PRECO", "ESTABELECIMENTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD USA_TAB_PRECO bit "

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select USUARIO.USUARIO_ID, USUARIO.SENHA, USUARIO.LOGON, USUARIO.TIPO, "
   SQL = SQL & " ESTABELECIMENTOACESSO.ESTABELECIMENTO_ID, ESTABELECIMENTO.DESCRICAO,"
   SQL = SQL & " ESTABELECIMENTO.Empresa_ID, ESTABELECIMENTO.CNPJCPF,CONTROLE_ESTOQUE, LOCALIZACAO, "
   SQL = SQL & " INDR_INDUSTRIA, LIBERA_DESCONTO,"
   SQL = SQL & " vlr_dia_compra_prod, USA_NFe, RECEBE_PEDIDO_VENDA, ESTOQUE_NEGATIVO,"
   SQL = SQL & " LEI_12741, cgc, razao_social,empresa.empresa_id,"
   SQL = SQL & " TamanhoCodgProdBarra,TamanhoPesoValorBarra,INDR_PANIFIC,peso_valor,CasaInicioCodgProdBarra,"
   SQL = SQL & " DESCONTO_CLIENTE , DESCONTO_FUNCIONARIO, AT_VENDA_MKP,versao_aplicativo,LiberaPercDesconto,"
   SQL = SQL & " DiasAtrazoCliente,CARTAOADM_ID,limpa_pedido,DOC_FISCAL,"
   SQL = SQL & " ALTERA_FATURA, USA_TAB_PRECO"

   SQL = SQL & " from USUARIO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ESTABELECIMENTOACESSO WITH (NOLOCK) "
   SQL = SQL & " ON USUARIO.USUARIO_ID = ESTABELECIMENTOACESSO.USUARIO_ID "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK) "
   SQL = SQL & " ON ESTABELECIMENTOACESSO.ESTABELECIMENTO_ID = ESTABELECIMENTO.ESTABELECIMENTO_ID "
   SQL = SQL & " INNER JOIN EMPRESA WITH (NOLOCK) "
   SQL = SQL & " ON ESTABELECIMENTO.EMPRESA_ID = EMPRESA.EMPRESA_ID"

   SQL = SQL & " where logon = '" & Trim(txtUsu.Text) & "'"
   SQL = SQL & " and senha = '" & (txtSenha.Text) & "'"
   SQL = SQL & " and ESTABELECIMENTOACESSO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabUSU.EOF Then
      MsgBox "Usuário inválido."
      Exit Sub
   End If

   USA_TAB_PRECO_B = False
   If Not IsNull(TabUSU.Fields("USA_TAB_PRECO").Value) Then _
      USA_TAB_PRECO_B = TabUSU.Fields("USA_TAB_PRECO").Value

   ALTERA_FATURA_B = False
   If Not IsNull(TabUSU.Fields("ALTERA_FATURA").Value) Then _
      ALTERA_FATURA_B = TabUSU.Fields("ALTERA_FATURA").Value

   VERSAO_APLICATIVO = "" & Trim(UCase(TabUSU.Fields("versao_aplicativo").Value))

   USU_LOGADO = Trim(txtUsu.Text)
   Valor_Compra_Dia_Permitida = 0

   Valor_Compra_Dia_Permitida = 0 & TabUSU.Fields("vlr_dia_compra_prod").Value
   'Variavel que checa se é pra bloquear compras para clientes apos dias de atrazo,
   'caso = 0 não rotina trata
   DiasAtrazoCliente_N = 0 & TabUSU.Fields("DiasAtrazoCliente").Value

   EMPRESA_ID_N = "" & TabUSU.Fields("empresa_id").Value
   CNPJ_EMPRESA_N = "" & Trim(TabUSU.Fields("cgc").Value)
   CNPJ_ESTABELECIMENTO_N = "" & Trim(TabUSU.Fields("CNPJCPF").Value)
   If Not IsNull(TabUSU.Fields("CARTAOADM_ID").Value) Then _
      CNPJ_CRED_CARTAO_ESTAB = "" & TRAZ_TEXTO_TABELA("CARTAOADM", "CNPJ", "CARTAOADM_ID", TabUSU.Fields("CARTAOADM_ID").Value)
   NOME_EMPRESA_A = "" & Trim(TabUSU.Fields("razao_social").Value)

   INDR_CONTROLA_ESTOQUE = False
   If Not IsNull(TabUSU.Fields("CONTROLE_ESTOQUE").Value) Then _
      INDR_CONTROLA_ESTOQUE = TabUSU.Fields("CONTROLE_ESTOQUE").Value

   INDR_AT_VENDA_MKP = False
   If Not IsNull(TabUSU!AT_VENDA_MKP) Then _
      INDR_AT_VENDA_MKP = TabUSU!AT_VENDA_MKP

   USA_NFe = False
   If Not IsNull(TabUSU!USA_NFe) Then _
      USA_NFe = TabUSU!USA_NFe

   USA_DOC_FISCAL = False
   If Not IsNull(TabUSU.Fields("DOC_FISCAL").Value) Then _
      USA_DOC_FISCAL = TabUSU.Fields("DOC_FISCAL").Value

   INDR_LEI_12741 = False
   If Not IsNull(TabUSU.Fields("LEI_12741").Value) Then _
      INDR_LEI_12741 = TabUSU.Fields("LEI_12741").Value

   INDR_INDUSTRIA_B = False
   If Not IsNull(TabUSU.Fields("INDR_INDUSTRIA").Value) Then _
      INDR_INDUSTRIA_B = TabUSU.Fields("INDR_INDUSTRIA").Value

   INDR_LIBERA_DESCONTO = False
   If Not IsNull(TabUSU.Fields("LIBERA_DESCONTO").Value) Then _
      INDR_LIBERA_DESCONTO = TabUSU.Fields("LIBERA_DESCONTO").Value

   INDR_DESCONTO_CLIENTE = False
   If Not IsNull(TabUSU.Fields("DESCONTO_CLIENTE").Value) Then _
      INDR_DESCONTO_CLIENTE = TabUSU.Fields("DESCONTO_CLIENTE").Value

   INDR_DESCONTO_FUNCIONARIO = False
   If Not IsNull(TabUSU.Fields("DESCONTO_FUNCIONARIO").Value) Then _
      INDR_DESCONTO_FUNCIONARIO = TabUSU.Fields("DESCONTO_FUNCIONARIO").Value

   INDR_LiberaPercDesconto = False
   If Not IsNull(TabUSU.Fields("LiberaPercDesconto").Value) Then _
      INDR_LiberaPercDesconto = TabUSU.Fields("LiberaPercDesconto").Value

   RECEBE_PEDIDO_VENDA = False
   If Not IsNull(TabUSU.Fields("RECEBE_PEDIDO_VENDA").Value) Then _
      RECEBE_PEDIDO_VENDA = TabUSU.Fields("RECEBE_PEDIDO_VENDA").Value

   LIMPA_PEDIDO = False
   If Not IsNull(TabUSU.Fields("limpa_pedido").Value) Then _
      LIMPA_PEDIDO = TabUSU.Fields("limpa_pedido").Value

   INDR_ESTQ_NEGATIVO = False
   If Not IsNull(TabUSU.Fields("ESTOQUE_NEGATIVO").Value) Then _
      INDR_ESTQ_NEGATIVO = TabUSU.Fields("ESTOQUE_NEGATIVO").Value

   '1=DICADO ; 2=IP ; 3=DEDICADO
   TIPO_TEF_N = 2

   '=====================
   '''''''''''balança
   CasaInicioCodgProdBarra_N = 0
   TamanhoCodgProdBarra_N = 0
   If Not IsNull(TabUSU.Fields("CasaInicioCodgProdBarra").Value) Then _
      CasaInicioCodgProdBarra_N = TabUSU.Fields("CasaInicioCodgProdBarra").Value

   If Not IsNull(TabUSU.Fields("TamanhoCodgProdBarra").Value) Then _
      TamanhoCodgProdBarra_N = TabUSU.Fields("TamanhoCodgProdBarra").Value

   TamanhoPesoValorBarra_N = 0
   If Not IsNull(TabUSU.Fields("TamanhoPesoValorBarra").Value) Then _
      TamanhoPesoValorBarra_N = TabUSU.Fields("TamanhoPesoValorBarra").Value

   If Not IsNull(TabUSU.Fields("INDR_PANIFIC").Value) Then
      If TabUSU.Fields("INDR_PANIFIC").Value = 0 Then
         MULT_EMPRESA_B = False
         Else: MULT_EMPRESA_B = True
      End If
   End If
   If Not IsNull(TabUSU.Fields("peso_valor").Value) Then _
      PESO_VALOR_A = Trim(UCase(TabUSU.Fields("peso_valor").Value))
   '''''''''''balança

   'If USA_NFC_E = True Then _
       Call frmDISPLAYEMISSOR.CarregarEasyTEF

   INDR_FIM = False
   Me.Hide
   frmINICIO.Show

Exit Sub
ERRO_TRATA:
   If Err.Number = 3704 Then _
      Resume Next
   TRATA_ERROS Err.Description, Me.Name, "SETA_BANCO"
End Sub

Sub CRIA_ACESSO()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTABELECIMENTOACESSO", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ESTABELECIMENTOACESSO]("
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL, [USUARIO_ID] Not [Int]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTABELECIMENTOACESSO]  WITH CHECK ADD  CONSTRAINT [FK_ESTABELECIMENTOACESSO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTABELECIMENTOACESSO] CHECK CONSTRAINT [FK_ESTABELECIMENTOACESSO_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTABELECIMENTOACESSO]  WITH CHECK ADD  CONSTRAINT [FK_ESTABELECIMENTOACESSO_USUARIO] FOREIGN KEY([USUARIO_ID])"
      SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTABELECIMENTOACESSO] CHECK CONSTRAINT [FK_ESTABELECIMENTOACESSO_USUARIO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
End Sub
