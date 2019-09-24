VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsultaPessoa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta"
   ClientHeight    =   6765
   ClientLeft      =   2175
   ClientTop       =   2565
   ClientWidth     =   10995
   Icon            =   "ConsultaPessoa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10995
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
      Left            =   9240
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtRazao 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1380
      Width           =   4695
   End
   Begin VB.CheckBox chkFunc 
      Caption         =   "&Funcionário"
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
      Left            =   9240
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox chkC 
      Caption         =   "&Cancelados"
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
      Left            =   9240
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtFone 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1380
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   0
      Top             =   900
      Width           =   4695
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   900
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##.###.###/####-##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
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
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7800
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaPessoa.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaPessoa.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaPessoa.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConsultaPessoa.frx":8F46
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   4845
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   8546
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CNPJ/CPF"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   6006
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Razão Social"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Cadastro"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "id"
         Object.Width           =   18
      EndProperty
   End
   Begin VB.Label lblRazao 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Razão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   525
      TabIndex        =   11
      Top             =   1440
      Width           =   570
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1095
      Left            =   0
      Top             =   720
      Width           =   10935
   End
   Begin VB.Label lblFone 
      Alignment       =   1  'Right Justify
      Caption         =   "Fone:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6525
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblCNPJCPF 
      Alignment       =   1  'Right Justify
      Caption         =   "CNPJ/CPF:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6090
      TabIndex        =   5
      Top             =   960
      Width           =   885
   End
   Begin VB.Label lblNome 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   4
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "frmConsultaPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "", "", "", ""

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
   PREPARA_TELA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_Load"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SETA_GRID
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Form_KeyPress"
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
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         SETA_GRID
      Case "limpar"
         NOME_A = ""
         FONE_A = ""
         CNPJCPF_A = ""
         txtNome.Text = ""
         txtRazao.Text = ""
         txtFone.Text = ""
         txtCNPJCPF.Text = ""
         chkC.Value = 0
         chkFunc.Value = 0
         LISTA.ListItems.Clear
         txtNome.SetFocus
      Case "voltar"
         Unload Me
      Case "print"
         FORMULA_REL = "{CLIENTE.nome} <> '' "

         If Trim(txtNome.Text) <> "" Then _
            FORMULA_REL = "{CLIENTE.nome} like '" & Trim(txtNome.Text) & "%'"

         If Trim(txtRazao.Text) <> "" Then _
            FORMULA_REL = "{CLIENTE.razao_social} like '" & Trim(txtRazao.Text) & "%'"

         If Trim(txtFone.Text) <> "" Then _
            FORMULA_REL = "{FONE.numero} like '" & Trim(txtFone.Text) & "%'"

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Cliente.rpt"
         frmRELATORIO10.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "Toolbar1_ButtonClick"
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Lista_DblClick()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = LISTA.SelectedItem.Text
   PESSOA_ID_N = LISTA.SelectedItem.ListSubItems.item(5)
   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "LISTA_DblClick"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe nome cliente e tecle <<ENTER>>", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtNome_GotFocus"
End Sub

Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe CNPJ/CPF cliente e tecle <<ENTER>>", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtCNPJCPF_GotFocus"
End Sub

Private Sub txtFone_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe Nº Telefone do cliente e tecle <<ENTER>>", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "txtfone_GotFocus"
End Sub
'================================
Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   HORA_INI = Time

   MOSTRA_RODAPE "Aguarde, Pesquisando ...", "", "", "", ""

   'ABRE_BANCO_SQLSERVER

   Dim RAZAO_A    As String
   Dim CNPCCPF_A  As String

   NOME_A = ""
   RAZAO_A = ""
   FONE_A = ""
   CNPJCPF_A = ""

   If Trim(txtCNPJCPF.Text) <> "" Then _
      CNPJCPF_A = "" & Chr$(39) & txtCNPJCPF.Text & "%" & Chr(39)
   If Trim(txtFone.Text) <> "" Then _
      FONE_A = "" & Chr$(39) & txtFone.Text & "%" & Chr(39)
   If Trim(txtNome.Text) <> "" Then _
      NOME_A = "" & Chr$(39) & txtNome.Text & "%" & Chr(39)
   If Trim(txtRazao.Text) <> "" Then _
      RAZAO_A = "" & Chr$(39) & txtRazao.Text & "%" & Chr(39)

   If Trim(TIPO_PESSOA_CADASTRO) = "CLIENTE" Then
      SQL = "select PESSOA.* FROM PESSOA "
      SQL = SQL & " INNER JOIN CLIENTE "
      SQL = SQL & " ON PESSOA.PESSOA_ID = CLIENTE.PESSOA_ID "
   End If
   If Trim(TIPO_PESSOA_CADASTRO) = "FORNECEDOR" Then
      SQL = "select PESSOA.* FROM PESSOA "
      SQL = SQL & " INNER JOIN FORNECEDOR "
      SQL = SQL & " ON PESSOA.PESSOA_ID = FORNECEDOR.PESSOA_ID "
   End If
   If Trim(TIPO_PESSOA_CADASTRO) = "USUARIO" Then
      SQL = "select PESSOA.* FROM PESSOA "
      SQL = SQL & " INNER JOIN USUARIO "
      SQL = SQL & " ON PESSOA.PESSOA_ID = USUARIO.PESSOA_ID "
   End If
   If Trim(TIPO_PESSOA_CADASTRO) = "TRANSPORTADORA" Then
      SQL = "select PESSOA.* FROM PESSOA "
      SQL = SQL & " INNER JOIN TRANSPORTADORA "
      SQL = SQL & " ON PESSOA.PESSOA_ID = TRANSPORTADORA.PESSOA_ID "
   End If

   SQL = SQL & " where PESSOA.pessoa_id is not null "

   If chkC.Value = 1 Then
      SQL = SQL & " and situacao = 'C'"
      Else: SQL = SQL & " and situacao = 'A'"
   End If

   If Trim(CNPJCPF_A) <> "" Then _
      SQL = SQL & " and cnpjcpf like " & CNPJCPF_A

   If Trim(NOME_A) <> "" Then _
      SQL = SQL & " and descricao LIKE " & NOME_A

   If Trim(RAZAO_A) <> "" Then _
      SQL = SQL & " and razao LIKE " & RAZAO_A

   If Trim(txtFone.Text) <> "" Then _
      SQL = SQL & " and numero like " & FONE_A

   SQL = SQL & " order by descricao "

   If TabTemp.State = 1 Then _
      TabTemp.Close

   LISTA.ListItems.Clear
   NUMR_SEQ_N = 0

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NUMR_SEQ_N = 1 + NUMR_SEQ_N

      Set item = LISTA.ListItems.Add(, "seq." & NUMR_SEQ_N, Trim(TabTemp!CNPJCPF))
      item.SubItems(1) = "" & Trim(TabTemp.Fields("descricao").Value)
      item.SubItems(2) = "" & Trim(TabTemp!RAZAO)
      item.SubItems(3) = "" & TabTemp!DaTa_CAD
      item.SubItems(4) = "" & TabTemp!SITUACAO
      item.SubItems(5) = "" & TabTemp.Fields("pessoa_id").Value
      TabTemp.MoveNext
   Wend
   TabTemp.Close

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

   HORA_FIM = Time

   MOSTRA_RODAPE "ESC - Sair", "Duplo click para selecionar", "Duração da consulta = " & Format((HORA_FIM - HORA_INI), "hh:mm:ss"), "Total de Registros Encontrados = " & NUMR_SEQ_N, ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "SETA_GRID"
End Sub

Sub PREPARA_TELA()
'On Error GoTo ERRO_TRATA

   lblNome.Visible = False
   txtNome.Visible = False
   lblCNPJCPF.Visible = False
   txtCNPJCPF.Visible = False
   chkFunc.Visible = False
   chkC.Visible = False
   lblRazao.Visible = False
   txtRazao.Visible = False

   If Trim(TIPO_PESSOA_CADASTRO) = "CLIENTE" Then
      Me.Caption = "Consulta Cliente"
      lblNome.Visible = True
      txtNome.Visible = True
      lblCNPJCPF.Visible = True
      txtCNPJCPF.Visible = True
      chkFunc.Visible = True
      chkC.Visible = True
      lblRazao.Visible = True
      txtRazao.Visible = True
   End If
   If Trim(TIPO_PESSOA_CADASTRO) = "FORNECEDOR" Then
      Me.Caption = "Consulta Fornecedor"
      lblNome.Visible = True
      txtNome.Visible = True
      lblCNPJCPF.Visible = True
      txtCNPJCPF.Visible = True
      chkC.Visible = True
      lblRazao.Visible = True
      txtRazao.Visible = True
   End If
   If Trim(TIPO_PESSOA_CADASTRO) = "USUARIO" Then
      Me.Caption = "Consulta Usuário"
      lblNome.Visible = True
      txtNome.Visible = True
      lblCNPJCPF.Visible = True
      txtCNPJCPF.Visible = True
      chkC.Visible = True
      chkFunc.Visible = True
   End If
   If Trim(TIPO_PESSOA_CADASTRO) = "TRANSPORTADORA" Then
      Me.Caption = "Consulta Transportadora"
      lblNome.Visible = True
      txtNome.Visible = True
      lblCNPJCPF.Visible = True
      txtCNPJCPF.Visible = True
      chkC.Visible = True
      lblRazao.Visible = True
      txtRazao.Visible = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "PREPARA_TELA"
End Sub
