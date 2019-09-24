VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROCONTA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contas Bancárias"
   ClientHeight    =   3300
   ClientLeft      =   3120
   ClientTop       =   3060
   ClientWidth     =   8865
   Icon            =   "cadastroconta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8865
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   50
      TabIndex        =   8
      Top             =   2040
      Width           =   8775
      Begin VB.TextBox txtConta 
         Alignment       =   2  'Center
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCli 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   9
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
      Begin MSMask.MaskEdBox txtCPF 
         Height          =   345
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##############"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Conta:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   255
         Width           =   840
      End
      Begin VB.Label lblTipoCli 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   3240
         TabIndex        =   10
         Top             =   255
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   50
      TabIndex        =   5
      Top             =   720
      Width           =   8775
      Begin VB.ComboBox cmbBancoAux 
         BackColor       =   &H80000004&
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
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbAgenciaAux 
         BackColor       =   &H80000004&
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
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbBanco 
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
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   6735
      End
      Begin VB.ComboBox cmbAgencia 
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
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Banco:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Agencia:"
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
         Left            =   480
         TabIndex        =   6
         Top             =   735
         Width           =   1035
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList3"
      DisabledImageList=   "ImageList3"
      HotImageList    =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Gravar Informações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Cadastro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "imp"
            Object.ToolTipText     =   "Impressão"
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
         Left            =   7440
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   6360
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
               Picture         =   "cadastroconta.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroconta.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroconta.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroconta.frx":90A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroconta.frx":A2D5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCADASTROCONTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim CodigoCliente As Long

Private Sub Form_Activate()
'On Error GoTo ERRO_TRATA

   cmbBanco.Clear
   cmbBancoAux.Clear

   If TabBANCO.State = 1 Then _
      TabBANCO.Close

   SQL = "select * from BANCO "
   SQL = SQL & " order by nome_banco "
   TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabBANCO.EOF
      cmbBancoAux.AddItem TabBANCO.Fields("banco_id").Value
      cmbBanco.AddItem Trim(TabBANCO!Nome_Banco) & " - " & TabBANCO!Codg_Banco
      TabBANCO.MoveNext
   Wend
   If TabBANCO.State = 1 Then _
      TabBANCO.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
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
      Case "imp"
         FORMULA_REL = ""
         If IsNumeric(cmbBancoAux.Text) Then
            FORMULA_REL = "{AGENCIA.banco} = '" & cmbBancoAux.Text & "'"
            If txtConta.Text <> "" Then _
               FORMULA_REL = FORMULA_REL & " and " & "{conta.numr_conta} = " & "'" & txtConta.Text & "'"
         End If
         
         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "relcontas.rpt"
         frmRELATORIO10.Show 1
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_CONTA
         SETA_GRID
      Case "gravar"
         txtCPF.PromptInclude = False
         If cmbAgenciaAUX.Text <> "" And _
            cmbBancoAux.Text <> "" And _
            Trim(txtCPF.Text) <> "" And _
            txtConta.Text <> "" And _
            txtDesc.Text <> "" Then
      
            GRAVA_DADOS
            LIMPA_CONTA
            SETA_GRID
         End If
      Case "print"
      Case "matar"
         If txtConta.Text <> "" Then
            Msg = "Confirma Exclusão da conta corrente?"
            PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then
               SQL = "delete  from CONTA "
               SQL = SQL & " where numr_conta='" & txtConta.Text & "'"
               CONECTA_RETAGUARDA.Execute SQL
               LIMPA_CONTA
               SETA_GRID
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmbBanco_Click()
'On Error GoTo ERRO_TRATA

   cmbBancoAux.ListIndex = cmbBanco.ListIndex
   If cmbBancoAux.Text <> "" Then
      SETA_GRID_B
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_Click"
End Sub

Private Sub cmbAgencia_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Selecione uma agência"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub cmbBanco_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Selecione um banco"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub cmbAgencia_Click()
'On Error GoTo ERRO_TRATA

   cmbAgenciaAUX.ListIndex = cmbAgencia.ListIndex

   If Trim(cmbAgencia.Text) <> "" Then _
      MOSTRA_AGENCIA

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbAgencia_Click"
End Sub

Private Sub cmbAgencia_LostFocus()
   If Trim(cmbAgencia.Text) <> "" Then _
      MOSTRA_AGENCIA
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCPF.PromptInclude = False

      If Trim(txtCPF.Text) <> "" Then
         If cmbAgenciaAUX.Text <> "" And cmbBancoAux.Text <> "" And txtConta.Text <> "" And txtDesc.Text <> "" Then

            GRAVA_DADOS

            txtCPF.PromptInclude = False
            txtCPF.Text = ""
            txtConta.Text = ""
            txtDesc.Text = ""
            txtCli.Text = ""

            SETA_GRID
         End If
      End If
      txtCPF.PromptInclude = True
      txtConta.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCPF_KeyPress"
End Sub

Private Sub cmbBanco_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(cmbBanco.Text) <> "" Then _
      MOSTRA_BANCO

   cmbBancoAux.ListIndex = cmbBanco.ListIndex
   If cmbBancoAux.Text <> "" Then
      cmbAgencia.Clear
      cmbAgenciaAUX.Clear

      If TabAGENCIA.State = 1 Then _
         TabAGENCIA.Close

      SQL = "select AGENCIA.* from AGENCIA "
      SQL = SQL & " INNER JOIN BANCO "
      SQL = SQL & " ON AGENCIA.BANCO_ID = BANCO.BANCO_ID"
      SQL = SQL & " and AGENCIA.banco_id = " & cmbBancoAux.Text
      SQL = SQL & "order by a.nome_agencia"
      TabAGENCIA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabAGENCIA.EOF
         cmbAgencia.AddItem TabAGENCIA!NUMR_AGENCIA & " - " & TabAGENCIA!nome_agencia
         cmbAgenciaAUX.AddItem TabAGENCIA.Fields("agencia_id").Value
         TabAGENCIA.MoveNext
      Wend
      If TabAGENCIA.State = 1 Then _
         TabAGENCIA.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_LostFocus"
End Sub

Private Sub cmbbanco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmbAgencia.SetFocus
   End If
End Sub

Private Sub cmbAgencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtConta.SetFocus
   End If
End Sub

Private Sub txtConta_GotFocus()
   MOSTRA_RODAPE "ESC - Sair", "Informe número da conta", "", "", ""
End Sub

Private Sub txtConta_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If cmbAgenciaAUX.Text <> "" And cmbBancoAux.Text <> "" And txtConta.Text <> "" Then
         If TabCONTA.State = 1 Then _
            TabCONTA.Close

         SQL = "select * from CONTA "
         SQL = SQL & " where codg_banco = '" & cmbBancoAux.Text & "'"
         SQL = SQL & " and numr_conta = '" & txtConta.Text & "'"
         SQL = SQL & " and numr_agencia = '" & cmbAgenciaAUX.Text & "'"
         TabCONTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCONTA.EOF Then
            KeyAscii = 0
            'txtSaldo.Text = TABCONTA!saldo

            If Not IsNull(TabCONTA!DESC_CONTA) Then _
               txtDesc.Text = TabCONTA!DESC_CONTA

            If Not IsNull(TabCONTA!PESSOA_ID) Then
               If TabPessoa.State = 1 Then _
                  TabPessoa.Close

               SQL = "select * from PESSOA"
               SQL = SQL & " where pessoa_id = " & TabCONTA.Fields("pessoa_ID").Value
               TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabPessoa.EOF Then
                  txtCPF.PromptInclude = False
                  txtCPF.Text = TabPessoa!CNPJCPF
                  txtCPF.PromptInclude = True
                  If Not IsNull(TabPessoa!DESCRICAO) Then _
                     txtCli.Text = TabPessoa!DESCRICAO
                  Else: MsgBox "Cliente/Fornecedor não cadastrado."
               End If
               If TabPessoa.State = 1 Then _
                  TabPessoa.Close
            End If
         End If
         If TabCONTA.State = 1 Then _
            TabCONTA.Close

         txtDesc.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtConta_KeyPress"
End Sub
      
Private Sub txtDesc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCPF.SetFocus
   End If
End Sub

Private Sub LIMPA_CONTA()
   txtCPF.PromptInclude = False
   txtCPF.Text = ""
   txtCli.Text = ""
   cmbAgenciaAUX.Text = ""
   cmbBancoAux.Text = ""
   cmbBanco.Text = ""
   cmbAgencia.Text = ""
   txtConta.Text = ""
   'txtSaldo.Text = ""
   txtDesc.Text = ""
   cmbBanco.SetFocus
End Sub

Private Sub SETA_GRID()
   SQL = "select * from CONTA c, BANCO b, AGENCIA a "
   SQL = SQL & " where c.numr_agencia=a.numr_agencia "
   SQL = SQL & "and a.banco=b.banco "
   SQL = SQL & "and c.numr_agencia='" & cmbAgenciaAUX.Text & "' "
   SQL = SQL & "and b.banco='" & cmbBancoAux.Text & "'"
End Sub

Private Sub SETA_GRID_B()
   SQL = "select * from CONTA c, BANCO b, AGENCIA a "
   SQL = SQL & " where c.numr_agencia=a.numr_agencia "
   SQL = SQL & "and a.banco=b.banco "
   'SQL = SQL & "and c.numr_agencia='" & cmbAgenciaaux.Text & "' "
   SQL = SQL & "and b.banco='" & cmbBancoAux.Text & "'"
End Sub

Sub GRAVA_DADOS()
   NOME_A = ""

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   SQL = "select * from PESSOA"
   SQL = SQL & " where cnpjcpf = '" & Trim(txtCPF.Text) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then
      NOME_A = TabPessoa.Fields("descricao").Value
      txtCli.Text = Trim(NOME_A)
      CodigoCliente = 0 & TabPessoa.Fields("pessoa_id").Value
      Else
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         MsgBox "CNPJ/CPF não encontrado"
         Exit Sub
   End If

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   If cmbAgenciaAUX.Text <> "" And cmbBancoAux.Text <> "" And txtConta.Text <> "" And txtDesc.Text <> "" Then
      If TabCONTA.State = 1 Then _
         TabCONTA.Close

      SQL = "select * from CONTA "
      SQL = SQL & " where numr_conta='" & txtConta.Text & "'"
      TabCONTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCONTA.EOF Then
         SQL = "UPDATE CONTA SET "
         SQL = SQL & " Banco = " & cmbBancoAux.Text
         SQL = SQL & ", NUMR_AGENCIA = '" & cmbAgenciaAUX.Text & "'"
         SQL = SQL & ", NUMR_CONTA = '" & txtConta.Text & "'"
         SQL = SQL & ", pessoa_id = " & CodigoCliente
         SQL = SQL & ", desc_conta = '" & txtDesc.Text & "'"
         SQL = SQL & " where numr_conta = '" & Trim(txtConta.Text) & "'"
         SQL = SQL & " and numr_agencia = '" & Trim(cmbAgenciaAUX.Text) & "'"
         Else
            SQL = "INSERT INTO CONTA "
            SQL = SQL & " (Banco, NUMR_AGENCIA, NUMR_CONTA, pessoa_id, desc_conta, dt_cadastro)"
            SQL = SQL & " VALUES ("
            SQL = SQL & cmbBancoAux.Text
            SQL = SQL & "," & cmbAgenciaAUX.Text
            SQL = SQL & ",'" & txtConta.Text & "'"
            SQL = SQL & "," & CodigoCliente
            SQL = SQL & ",'" & txtDesc.Text & "'"
            SQL = SQL & ",'" & Now & "'"
            SQL = SQL & ")"
      End If
      If TabCONTA.State = 1 Then _
         TabCONTA.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If

End Sub


Sub MOSTRA_BANCO()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(cmbBanco.Text) <> "" Then
      SQL = "select * from Banco "

      If IsNumeric(cmbBanco.Text) Then
         SQL = SQL & " where codg_banco = '" & cmbBanco.Text & "'"
         Else: SQL = SQL & " where nome_banco = '" & Trim(cmbBanco.Text) & "'"
      End If

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         'lblBanco.Caption = Trim(TABTEMP!Nome_Banco) & " - " & TABTEMP.Fields("banco").Value
         cmbBanco.Text = Trim(TabTemp!Nome_Banco)
         cmbBancoAux.Text = Trim(TabTemp.Fields("banco").Value)
      End If
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_BANCO"
End Sub

Sub MOSTRA_AGENCIA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(cmbAgencia.Text) <> "" Then
      SQL = "select * from AGENCIA "

      If IsNumeric(cmbAgencia.Text) Then
         SQL = SQL & " where numr_AGENCIA = '" & cmbAgencia.Text & "'"
         Else: SQL = SQL & " where nome_agencia = '" & Trim(cmbAgencia.Text) & "'"
      End If

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         'lblAgencia.Caption = Trim(TABTEMP!NOME_AGENCIA) & " - " & TABTEMP.Fields("numr_AGENCIA").Value
         cmbAgencia.Text = TabTemp.Fields("numr_AGENCIA").Value & "-" & Trim(TabTemp!nome_agencia)
         cmbAgenciaAUX.Text = TabTemp.Fields("numr_AGENCIA").Value
      End If
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_AGENCIA"
End Sub
