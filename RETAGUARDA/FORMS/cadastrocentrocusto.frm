VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROCENTROCUSTO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Centro de Custo"
   ClientHeight    =   4920
   ClientLeft      =   3075
   ClientTop       =   2895
   ClientWidth     =   9945
   Icon            =   "cadastrocentrocusto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9945
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   50
      TabIndex        =   6
      Top             =   30
      Width           =   8895
      Begin VB.ComboBox cmbAuxCC 
         BackColor       =   &H00FFC0C0&
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
         Left            =   2160
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
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
         Left            =   7200
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtNome 
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
         Height          =   360
         Left            =   3435
         MaxLength       =   50
         TabIndex        =   11
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtCodg 
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
         Height          =   360
         Left            =   2115
         MaxLength       =   20
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbCC 
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
         Left            =   2115
         TabIndex        =   9
         Top             =   720
         Width           =   4935
      End
      Begin VB.OptionButton optD 
         Caption         =   "&Débito"
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
         Left            =   7155
         TabIndex        =   8
         Top             =   900
         Width           =   1335
      End
      Begin VB.OptionButton optC 
         Caption         =   "&Crédito"
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
         Left            =   7155
         TabIndex        =   7
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo:"
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
         Left            =   435
         TabIndex        =   13
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grupo Centro Custo:"
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
         Left            =   90
         TabIndex        =   12
         Top             =   765
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Left            =   9000
      Picture         =   "cadastrocentrocusto.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   900
   End
   Begin VB.CommandButton cmbVoltar 
      Caption         =   "&Voltar"
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
      Left            =   9000
      Picture         =   "cadastrocentrocusto.frx":6BB7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Enabled         =   0   'False
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
      Left            =   9000
      Picture         =   "cadastrocentrocusto.frx":7D41
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   900
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
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
      Left            =   9000
      Picture         =   "cadastrocentrocusto.frx":8F99
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   900
   End
   Begin VB.CommandButton cmdMatar 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
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
      Left            =   9000
      Picture         =   "cadastrocentrocusto.frx":A018
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   900
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   3495
      Left            =   45
      TabIndex        =   5
      Top             =   1320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codg."
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
         Text            =   "Tipo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Grp."
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmCADASTROCENTROCUSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcc_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Selecione um grupo do CCusto"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub cmdImprimir_Click()
FORMULA_REL = ""

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "relccusto.rpt"
   frmRELATORIO10.Show 1
End Sub

Private Sub Form_Load()
   Call CentralizaJanela2(frmCADASTROCENTROCUSTO)
End Sub

Private Sub Form_Resize()
   LISTA.ListItems.Clear
   SQL = "select * from CCUSTO c "
   SQL = SQL & " left join DESCR d "
   SQL = SQL & " on c.grupo_cc = d.codigo "
   SQL = SQL & " where d.TIPO = 'O' "
   SQL = SQL & "order by c.DESCR_CC asc"
   SP_SETA_CCUSTO
   SP_LIMPA_CCUSTO

   cmbAuxCC.Clear
   cmbCC.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close
   SQL = "select * from DESCR "
   SQL = SQL & " where TIPO = 'O'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      cmbCC.AddItem Trim(TabTemp!DESCRICAO)
      cmbAuxCC.AddItem TabTemp!Codigo
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         SP_LIMPA_CCUSTO
         txtCodg.SetFocus
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If txtCodg.Text <> "" Then
      CmdGravar.Enabled = True
      cmdMatar.Enabled = True
      Else
         CmdGravar.Enabled = False
         cmdMatar.Enabled = False
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

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub optC_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Seleciona a opção crédito"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub optD_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Seleciona a opção débito"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub txtCodg_Change()
   If txtCodg.Text <> "" Then
      CRITERIO_A = Chr$(39) & txtCodg.Text & "%" & Chr(39)

      LISTA.ListItems.Clear
      SQL = "select * from CCUSTO c "
      SQL = SQL & " inner join DESCR d "
      SQL = SQL & " on c.grupo_cc = d.codigo "
      SQL = SQL & " where d.TIPO = 'O' "
      SQL = SQL & " and c.codg_cc like " & CRITERIO_A
      SQL = SQL & "order by c.DESCR_CC asc"
      SP_SETA_CCUSTO
   End If
End Sub

Private Sub txtCodg_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o código do CCusto"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub txtcodg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtCodg.Text = "" Then
         txtCodg.Text = 1
         SQL = "select max(codg_cc) as ultimo_rg from CCUSTO "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            If Not IsNull(TabTemp!ultimo_rg) Then _
               txtCodg.Text = TabTemp!ultimo_rg + 1
         End If
      End If
      CRITERIO_A = txtCodg.Text
      SP_PROCURA_CCUSTO
      If Not TabTemp.EOF Then
         KeyAscii = 0
         txtNome.Text = TabTemp!descr_cc
         If TabTemp!tipo_cc <> "" Then
            If Not IsNull(TabTemp!tipo_cc) Then
               If TabTemp!tipo_cc = "D" Then
                  optD.Value = True
               End If
               If TabTemp!tipo_cc = "C" Then
                  optC.Value = True
               End If
            End If
         End If
         If Not IsNull(TabTemp!grupo_cc) Then
            If Not TabTemp!grupo_cc = "" Then
               SQL = "select * from DESCR "
               SQL = SQL & " where TIPO = 'O' "
               SQL = SQL & " and codigo = '" & Trim(TabTemp!grupo_cc) & "'"
               TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabTemp.EOF Then
                  cmbCC.Text = Trim(TabTemp!DESCRICAO)
                  cmbAuxCC.Text = TabTemp!Codigo
               End If
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      txtNome.SetFocus
   End If
End Sub

Private Sub txtNome_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o nome do CCusto"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txtNome.Text) = "" Then
         MsgBox "Informe Nome do CCUSTO !!!"
         txtNome.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      txtCodg.SetFocus
   End If
End Sub

Private Sub txtNome_LostFocus()
   txtNome.Text = UCase(txtNome)
End Sub

Private Sub cmbVoltar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   SP_LIMPA_CCUSTO
   LISTA.ListItems.Clear
   SQL = "select * from CCUSTO c "
   SQL = SQL & " inner join DESCR d "
   SQL = SQL & " on c.grupo_cc = d.codigo "
   SQL = SQL & " where d.TIPO = 'O' "
   SQL = SQL & "order by c.DESCR_CC asc"
   SP_SETA_CCUSTO
   txtCodg.SetFocus
End Sub

Private Sub CmdGravar_Click()
   If txtCodg.Text <> "" And Trim(txtNome.Text) <> "" Then
      If cmbAuxCC.Text = "" Then cmbAuxCC.Text = 0
      CRITERIO_A = txtCodg.Text
      SP_PROCURA_CCUSTO
      If Not TabTemp.EOF Then
         SQL = "update CCUSTO set "
         SQL = SQL & "CODG_CC = " & txtCodg.Text
         SQL = SQL & ",GRUPO_CC = " & cmbAuxCC.Text
         SQL = SQL & ",DESCR_CC='" & txtNome.Text & "'"
         If optD.Value = True Then
            SQL = SQL & ",TIPO_CC='D'"
            Else: SQL = SQL & ",TIPO_CC='O'"
         End If
         SQL = SQL & " where CODG_CC = " & txtCodg.Text
         CONECTA_RETAGUARDA.Execute SQL
      Else
         SqL2 = "INSERT INTO CCUSTO (Codg_cc, descr_cc, grupo_cc, Tipo_cc)"
         SqL2 = SqL2 & " VALUES (" & txtCodg.Text & ",'" & txtNome.Text & "'," & cmbAuxCC.Text & ",'" & "D" & "')"
         CONECTA_RETAGUARDA.Execute SqL2
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      LISTA.ListItems.Clear
      SQL = "select * from CCUSTO c "
      SQL = SQL & " inner join DESCR d "
      SQL = SQL & " on c.grupo_cc = d.codigo "
      SQL = SQL & " where d.TIPO = 'O' "
      SQL = SQL & "order by c.DESCR_CC asc"
      SP_SETA_CCUSTO
      SP_LIMPA_CCUSTO
      txtCodg.SetFocus
      Else: MsgBox "Informe os dados corretamente !!!"
   End If
End Sub

Private Sub cmdMatar_Click()
   If txtCodg.Text <> "" And Trim(txtNome.Text) <> "" Then
      CRITERIO_A = txtCodg.Text
      SP_PROCURA_CCUSTO

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SP_LIMPA_CCUSTO

      LISTA.ListItems.Clear

      SQL = "select * from CCUSTO c "
      SQL = SQL & " inner join DESCR d "
      SQL = SQL & " on c.grupo_cc = d.codigo "
      SQL = SQL & " where d.TIPO = 'O' "
      SQL = SQL & "order by c.DESCR_CC asc"
      SP_SETA_CCUSTO
      txtCodg.SetFocus
   End If
End Sub

Private Sub cmbcc_LostFocus()
   cmbAuxCC.ListIndex = cmbCC.ListIndex
End Sub
'==========================================
Private Sub SP_LIMPA_CCUSTO()
   txtCodg.Text = ""
   txtNome.Text = ""
   cmbAuxCC.Text = ""
   cmbCC.Text = ""
   optD.Value = True
   optC.Value = False
   CmdGravar.Enabled = False
   cmdMatar.Enabled = False
End Sub

Private Sub SP_SETA_CCUSTO()
   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set ITEM2 = LISTA.ListItems.Add(, "seq." & TabTemp!Codg_cc, TabTemp!Codg_cc)
      ITEM2.SubItems(1) = TabTemp!descr_cc
      ITEM2.SubItems(2) = TabTemp!tipo_cc
      ITEM2.SubItems(3) = Trim(TabTemp!DESCRICAO)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
End Sub

Public Sub SP_PROCURA_CCUSTO()
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from CCUSTO "
   SQL = SQL & " where CODG_CC = " & CRITERIO_A
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub
