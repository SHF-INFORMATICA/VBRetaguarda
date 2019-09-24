VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmControleAcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Acesso"
   ClientHeight    =   7455
   ClientLeft      =   1620
   ClientTop       =   2340
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ControleAcesso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8445
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Controle de Acesso"
      ForeColor       =   &H00400000&
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Voltar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   7320
         Picture         =   "ControleAcesso.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1020
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   5160
         Picture         =   "ControleAcesso.frx":6D9C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   7440
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdGravar 
         Caption         =   "Gravar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   6240
         MaskColor       =   &H00FF8080&
         Picture         =   "ControleAcesso.frx":739E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Confirma os acessos para este usuario."
         Top             =   240
         Width           =   1005
      End
      Begin VB.ComboBox cmbUsuAux 
         BackColor       =   &H80000000&
         Height          =   405
         Left            =   1440
         TabIndex        =   9
         ToolTipText     =   "Selecione um usuario."
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame fraCopiaPerfil 
         Caption         =   "Copiar Perfil"
         ForeColor       =   &H00400000&
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   5880
         Width           =   8295
         Begin VB.ComboBox cmbUsuDestinoAUX 
            BackColor       =   &H80000000&
            Height          =   405
            Left            =   2280
            TabIndex        =   19
            ToolTipText     =   "Selecione um usuario."
            Top             =   960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbUsuDestino 
            Height          =   405
            Left            =   2280
            TabIndex        =   18
            ToolTipText     =   "Selecione um usuario, que receberá os acessos."
            Top             =   960
            Width           =   4335
         End
         Begin VB.ComboBox cmbUsuOrigemAUX 
            BackColor       =   &H80000000&
            Height          =   405
            Left            =   2280
            TabIndex        =   16
            ToolTipText     =   "Selecione um usuario."
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbUsuOrigem 
            Height          =   405
            Left            =   2280
            TabIndex        =   7
            ToolTipText     =   "Selecione um usuario, que receberá os acessos."
            Top             =   480
            Width           =   4335
         End
         Begin VB.CommandButton cmdCopiaPerfil 
            Caption         =   "Copiar Perfil"
            Height          =   855
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Clique para copiar o perfil para este usuario."
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Usuário Destino:"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Usuário Origem:"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbUsu 
         Height          =   405
         Left            =   1320
         TabIndex        =   1
         ToolTipText     =   "Selecione um usuario."
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame Frame2 
         ForeColor       =   &H00400000&
         Height          =   4815
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   8295
         Begin VB.CheckBox chkTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   3000
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin MSComctlLib.ListView LstOpcoes 
            Height          =   4095
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Clique para selecionar"
            Top             =   600
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   7223
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
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
               Text            =   "Rotina"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   8819
            EndProperty
         End
         Begin VB.Label Label3 
            Caption         =   "Opções do Sistema"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2340
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   990
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   8445
      DesignHeight    =   7455
   End
End
Attribute VB_Name = "FrmControleAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call CentralizaJanela(Me)

   cmbUsu.Clear

   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   'AbreBD
   PreencheComboUsuario

   PreencheComboUsuarioPerfil
End Sub

Private Sub Form_Terminate()
    If Len(Trim(cmbUsu.Text)) > 0 Then
        If MsgBox("As alterações efetuadas só serão efetivadas se o botão 'GRAVAR' tiver sido pressionado" & vbCrLf & "Deseja sair ?", vbInformation + vbYesNo + vbDefaultButton3, "Sair") <> vbYes Then
            CmdGravar.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub cmdLimpar_Click()
   cmbUsu.Text = ""
   cmbUsuAux.Text = ""
   chkTodos.Value = 0
   LstOpcoes.ListItems.Clear
   cmbUsuOrigem.Text = ""
   cmbUsuOrigemAUX.Text = ""
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmbUsu_Click()
'On Error GoTo ERRO_TRATA

   Dim IntInicio As Integer
   Dim StrCodUsuario As String

   If Trim(cmbUsu.Text) > 0 Then
      cmbUsuAux.ListIndex = cmbUsu.ListIndex
      StrCodUsuario = cmbUsuAux.Text

      PREENCHE_LISTA_MENU

      HabilitaFrames

      LstOpcoes.SetFocus
   End If

Err.Clear

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "cmbUsu_Click"
End Sub

Private Sub cmbUsu_LostFocus()
On Error Resume Next

   Dim IntInicio As Integer
   Dim StrCodUsuario As String

   If Len(cmbUsu.Text) > 0 Then
      IntInicio = InStr(1, Trim(cmbUsu.Text), "-")
      StrCodUsuario = Mid(Trim(cmbUsu.Text), (IntInicio + 1))
      PREENCHE_LISTA_MENU
      LstOpcoes.SetFocus
      Label4.Caption = "Acesso para"
      Label4.Caption = Label4.Caption & " " & cmbUsu.Text
      
      HabilitaFrames
      LstOpcoes.SetFocus
   End If
End Sub

Private Sub cmbUSUORIGEM_Click()
On Error Resume Next

   cmbUsuOrigemAUX.ListIndex = cmbUsuOrigem.ListIndex

Err.Clear
End Sub

Private Sub cmbUSUDESTINO_Click()
On Error Resume Next

   cmbUsuDestinoAUX.ListIndex = cmbUsuDestino.ListIndex

Err.Clear
End Sub

Private Sub CmdCopiaPerfil_Click()
'On Error GoTo ERRO_TRATA

   If Trim(cmbUsuDestinoAUX.Text) = "" Then
      MsgBox "Informar usuário destino."
      Exit Sub
   End If
   If Trim(cmbUsuOrigemAUX.Text) = "" Then
      MsgBox "Informar usuário origem."
      Exit Sub
   End If

   Dim USER_ORIGEM_N    As Long
   Dim USER_DESTINO_N   As Long

   USER_ORIGEM_N = cmbUsuOrigemAUX.Text
   USER_DESTINO_N = cmbUsuDestinoAUX.Text

   SQL = "Delete from PERMISSAO Where USUID = " & USER_DESTINO_N
   CONECTA_RETAGUARDA.Execute SQL

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from PERMISSAO WITH (NOLOCK)"
   SQL = SQL & " where usuid = " & USER_ORIGEM_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      SQL = "Insert Into PERMISSAO "
         SQL = SQL & " (MENUID, USUID, ACESSO) "
      SQL = SQL & " Values ("
         SQL = SQL & "'" & Trim(TabConsulta.Fields("Menuid").Value) & "'"
         SQL = SQL & "," & USER_DESTINO_N
         SQL = SQL & ",1"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   GRAVA_ESTABELECIMENTOACESSO USER_ORIGEM_N, ESTABELECIMENTO_ID_N

   MsgBox "Perfil copiado do usuario : " & cmbUsu.Text & " para usuario : " & cmbUsuOrigem & " com sucesso.", vbInformation, "MEGASIM"

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CmdCopiaPerfil_Click"
End Sub

Private Sub CmdGravar_Click()
'On Error GoTo ERRO_TRATA

   Dim IntInicio As Integer
   Dim StrCodUsuario As Long
   Dim strSQL As String
   Dim intLixo As Integer
   Dim i                   As Integer

   If LstOpcoes.ListItems.Count > 0 And Trim(cmbUsu.Text) > 0 Then
      strSQL = "Delete from PERMISSAO Where USUID = " & cmbUsuAux.Text
      CONECTA_RETAGUARDA.Execute (strSQL)

      For i = LstOpcoes.ListItems.Count To 1 Step -1
         If LstOpcoes.ListItems(i).Checked = True Then
            If Trim(UCase(LstOpcoes.ListItems(i).SubItems(1))) <> "" Then
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               SQL = "select * from PERMISSAO WITH (NOLOCK)"
               SQL = SQL & " where menuid = '" & Trim(LstOpcoes.ListItems(i).SubItems(1)) & "'"
               SQL = SQL & " and usuid = '" & Trim(cmbUsuAux.Text) & "'"
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabConsulta.EOF Then
                  If TabConsulta.State = 1 Then _
                     TabConsulta.Close

                  SQL = "Insert Into PERMISSAO (MENUID, USUID, Acesso) "
                  SQL = SQL & " Values ("
                  SQL = SQL & "'" & Trim(LstOpcoes.ListItems(i).SubItems(1)) & "'"
                  SQL = SQL & "," & Trim(cmbUsuAux.Text)
                  SQL = SQL & ",'true'"
                  SQL = SQL & ")"
                  CONECTA_RETAGUARDA.Execute (SQL)
               End If
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               DoEvents
            End If
         End If
      Next i
      'ESTABELECIMENTOACESSO
      GRAVA_ESTABELECIMENTOACESSO cmbUsuAux.Text, ESTABELECIMENTO_ID_N

      MsgBox "Processo realizado com sucesso !!!"
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "CmdGravar_Click"
End Sub

Private Sub chkTodos_Click()
'On Error GoTo ERRO_TRATA

   Dim i

   If LstOpcoes.ListItems.Count > 0 Then
      For i = LstOpcoes.ListItems.Count To 1 Step -1
         If chkTodos.Value = 1 Then
            LstOpcoes.ListItems(i).Checked = True
            Else: LstOpcoes.ListItems(i).Checked = False
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "chkTodos_Click"
End Sub

Private Sub LstOpcoes_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyDelete
         If Not IsNull(LstOpcoes.SelectedItem.ListSubItems.item(1).Text) Then
            Dim PAI_A   As String
            Dim PAI_id  As String

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "select MENUID, DESCMENU from MENU WITH (NOLOCK)"
            SQL = SQL & " where menuid = '" & Trim(LstOpcoes.SelectedItem.ListSubItems.item(1).Text) & "'"
            TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabTemp.EOF Then
               'achei o pai
               PAI_A = Trim(TabTemp!descmenu)
               PAI_id = Trim(TabTemp!menuid)

               SQL = "delete from PERMISSAO where menuid = '" & Trim(PAI_id) & "'"
               CONECTA_RETAGUARDA.Execute SQL

               SQL = "delete from MENU where menuid = '" & Trim(PAI_id) & "'"
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabTemp.State = 1 Then _
               TabTemp.Close
         End If
   End Select

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "LstOpcoes_KeyDown"
End Sub

Private Sub Command1_Click()
SQL = "delete from PERMISSAO where left(PERMISSAO.menuid,3) = 'frm'"
End Sub

Private Sub HabilitaFrames()
   Frame2.Enabled = True ' Frame com opcoes de menu
   fraCopiaPerfil.Enabled = True ' Frame com perfil copiar
End Sub

Private Sub PREENCHE_LISTA_MENU()
'On Error GoTo ERRO_TRATA

   LstOpcoes.ListItems.Clear
   CONT_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close
    
   SQL = "select MENUID, DESCMENU from MENU WITH (NOLOCK)"
   SQL = SQL & " order by DESCMENU"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   TabTemp.MoveFirst
   Do Until TabTemp.EOF
      CONT_N = CONT_N + 1
      Set item = LstOpcoes.ListItems.Add(, "seq." & CONT_N, Trim(TabTemp!descmenu))
      item.SubItems(1) = "" & Trim(TabTemp.Fields("menuid").Value)

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from PERMISSAO WITH (NOLOCK)"
      SQL = SQL & " where menuid = '" & Trim(TabTemp.Fields("menuid").Value) & "'"
      SQL = SQL & " and usuid = '" & Trim(cmbUsuAux.Text) & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         'Item.ListItems.Checked = True
         LstOpcoes.ListItems(CONT_N).Checked = True
         'Item.SubItems.Checked = True
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      TabTemp.MoveNext
   Loop
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "PREENCHE_LISTA_MENU"
End Sub

Private Sub PreencheComboUsuario()
'On Error GoTo ERRO_TRATA

   cmbUsu.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select EMPRESA_ID, usuario_id, NOME from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where status = 1"
   SQL = SQL & " Order by USUARIO.NOME"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   Do Until TabTemp.EOF
      cmbUsu.AddItem Trim(TabTemp!NOME) & "-" & Trim(TabTemp!USUARIO_ID)
      cmbUsuAux.AddItem Trim(TabTemp!USUARIO_ID)

      TabTemp.MoveNext
   Loop
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "PreencheComboUsuario"
End Sub

Private Sub PreencheComboUsuarioPerfil()
'On Error GoTo ERRO_TRATA

   cmbUsuOrigem.Clear
   cmbUsuOrigemAUX.Clear
   cmbUsuDestino.Clear
   cmbUsuDestinoAUX.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select EMPRESA_ID, usuario_id, NOME from USUARIO"
   SQL = SQL & " Order by USUARIO.NOME"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   Do Until TabTemp.EOF
      cmbUsuOrigem.AddItem Trim(TabTemp!NOME) & "-" & Trim(TabTemp!USUARIO_ID)
      cmbUsuOrigemAUX.AddItem Trim(TabTemp!USUARIO_ID)

      cmbUsuDestino.AddItem Trim(TabTemp!NOME) & "-" & Trim(TabTemp!USUARIO_ID)
      cmbUsuDestinoAUX.AddItem Trim(TabTemp!USUARIO_ID)

      TabTemp.MoveNext
   Loop
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "PreencheComboUsuarioPerfil"
End Sub
