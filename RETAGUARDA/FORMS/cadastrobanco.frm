VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROBANCO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Banco"
   ClientHeight    =   5505
   ClientLeft      =   1545
   ClientTop       =   930
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "cadastrobanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   50
      TabIndex        =   4
      Top             =   720
      Width           =   8775
      Begin VB.TextBox txtCodg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   5760
      End
      Begin MSComctlLib.ListView LISTA 
         Height          =   3825
         Left            =   80
         TabIndex        =   2
         Top             =   720
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   6747
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Banco:"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   1170
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastrobanco.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastrobanco.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastrobanco.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastrobanco.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastrobanco.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastrobanco.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastrobanco.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Description     =   "Voltar para Tela Início"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   6
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
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Cadastro"
            ImageIndex      =   7
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
         Left            =   7560
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrobanco.frx":76BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrobanco.frx":8AE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrobanco.frx":9B75
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrobanco.frx":ADDD
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrobanco.frx":BD92
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrobanco.frx":D48F
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastrobanco.frx":E629
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCADASTROBANCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela2(frmCADASTROBANCO)

   'ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         LIMPA_BANCO
         txtCodg.SetFocus
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

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "matar"
         If txtCodg.Text <> "" Then
            If IsNumeric(txtCodg.Text) Then
               If TabBANCO.State = 1 Then _
                  TabBANCO.Close

               SP_PROC_BANCO txtCodg.Text
               If Not TabBANCO.EOF Then
                  If TabBANCO.State = 1 Then _
                     TabBANCO.Close

                  SQL = "delete from BANCO "
                  SQL = SQL & " where codg_banco = '" & txtCodg.Text & "'"
                  CONECTA_RETAGUARDA.Execute SQL

                  LIMPA_BANCO
                  SETA_GRID
                  txtCodg.SetFocus
               End If
               If TabBANCO.State = 1 Then _
                  TabBANCO.Close
            End If
         End If
      Case "limpar"
         LIMPA_BANCO
         txtCodg.SetFocus
      Case "gravar"
         GRAVA_BANCO
      Case "print"
         FORMULA_REL = ""

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "relbanco.rpt"
         frmRELATORIO10.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCodg_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe código do banco", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodg_GotFocus"
End Sub

Private Sub txtcodg_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If txtCodg.Text = "" Then
         MsgBox "Informe Código do Banco !!!"
         txtCodg.SetFocus
         Exit Sub
      End If
      If txtCodg.Text <> "" Then
         If IsNumeric(txtCodg.Text) Then
            SP_PROC_BANCO txtCodg.Text
            If Not TabBANCO.EOF Then
               KeyAscii = 0
               txtNome.Text = TabBANCO!Nome_Banco
            End If
            If TabBANCO.State = 1 Then _
               TabBANCO.Close
         End If
      End If
      txtNome.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodg_KeyPress"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe nome do banco", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txtNome.Text) = "" Then
         MsgBox "Informe Nome do Banco !!!"
         txtNome.SetFocus
         Exit Sub
      End If
      KeyAscii = 0

      GRAVA_BANCO

      txtCodg.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyPress"
End Sub
'==========================================
Private Sub LIMPA_BANCO()
'On Error GoTo ERRO_TRATA

   txtCodg.Text = ""
   txtNome.Text = ""

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BANCO"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   NUMR_SEQ_N = 0
   LISTA.ListItems.Clear

   If TabBANCO.State = 1 Then _
      TabBANCO.Close

   SQL = "select * from BANCO "
   SQL = SQL & "order by banco"
   TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabBANCO.EOF
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set item = LISTA.ListItems.Add(, "seq." & NUMR_SEQ_N, TabBANCO!Codg_Banco)
      item.SubItems(1) = TabBANCO!Nome_Banco
      TabBANCO.MoveNext
   Wend
   If TabBANCO.State = 1 Then _
      TabBANCO.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub GRAVA_BANCO()
'On Error GoTo ERRO_TRATA

   If txtCodg.Text <> "" And Trim(txtNome.Text) <> "" Then
      If txtCodg.Text <> "" Then
         If TabBANCO.State = 1 Then _
            TabBANCO.Close

         SP_PROC_BANCO txtCodg.Text

         If Not TabBANCO.EOF Then
            SQL = "update BANCO set "
            SQL = SQL & "banco = '" & Trim(txtCodg.Text) & "'"
            SQL = SQL & ",nome_banco = '" & Trim(txtNome.Text) & "' "
            SQL = SQL & " where codg_banco = '" & Trim(txtCodg.Text) & "'"
            Else
               SQL = "INSERT INTO BANCO "
                  SQL = SQL & " (banco, Nome_banco)"
               SQL = SQL & " VALUES ("
               SQL = SQL & "'" & Trim(txtCodg.Text) & "'"
               SQL = SQL & ",'" & Trim(txtNome.Text) & "')"
         End If
         If TabBANCO.State = 1 Then _
            TabBANCO.Close

         CONECTA_RETAGUARDA.Execute SQL
      End If

      SETA_GRID
      LIMPA_BANCO

      txtCodg.SetFocus
      Else: MsgBox "Informe os dados corretamente !!!"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_BANCO"
End Sub
