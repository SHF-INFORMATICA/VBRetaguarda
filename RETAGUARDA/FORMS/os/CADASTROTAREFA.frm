VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmCADASTROTAREFA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Tarefa"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   960
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CADASTROTAREFA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   45
      TabIndex        =   8
      Top             =   720
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro Tarefa/Serviço"
      TabPicture(0)   =   "CADASTROTAREFA.frx":47C4A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Listagem Tarefa/Serviço"
      TabPicture(1)   =   "CADASTROTAREFA.frx":47C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LISTATAREFA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtDesc2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtDesc2 
         Height          =   360
         Left            =   -72300
         MaxLength       =   100
         TabIndex        =   5
         ToolTipText     =   "Informe "
         Top             =   480
         Width           =   5535
      End
      Begin VB.Frame Frame3 
         Height          =   2175
         Left            =   50
         TabIndex        =   9
         Top             =   600
         Width           =   11535
         Begin VB.TextBox txtCodigo 
            Alignment       =   2  'Center
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
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   0
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtDescricao 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   3600
            MultiLine       =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   7815
         End
         Begin VB.TextBox txtPerc 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   3
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   2
            Top             =   1680
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txtDtCad 
            Height          =   375
            Left            =   8400
            TabIndex        =   4
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Cadastro:"
            Height          =   240
            Left            =   6960
            TabIndex        =   14
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%Comissão = "
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
            Left            =   3750
            TabIndex        =   13
            Top             =   1680
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tarefa/Serviço:"
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
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor = "
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
            Left            =   1065
            TabIndex        =   10
            Top             =   1680
            Width           =   750
         End
      End
      Begin MSComctlLib.ListView LISTATAREFA 
         Height          =   3705
         Left            =   -74940
         TabIndex        =   6
         Top             =   960
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   6535
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   12582912
         BackColor       =   16777215
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "req"
            Text            =   "Código"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "cli"
            Text            =   "Descrição"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "valor"
            Text            =   "Valor Serviço"
            Object.Width           =   2382
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Key             =   "desconto"
            Text            =   "% Comissão"
            Object.Width           =   2294
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição Tarefa/Serviço:"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -74805
         TabIndex        =   12
         Top             =   480
         Width           =   2445
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "CADASTROTAREFA.frx":47C82
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAREFA.frx":480D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAREFA.frx":483F2
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAREFA.frx":48846
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAREFA.frx":48C9A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAREFA.frx":48FBA
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CADASTROTAREFA.frx":4940E
            Key             =   "IMG7"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barTAREFA 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   1111
      ButtonWidth     =   1191
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageKey        =   "IMG7"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "matar"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Observações"
            Key             =   "obs"
            ImageKey        =   "IMG5"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11745
      DesignHeight    =   5550
   End
End
Attribute VB_Name = "frmCADASTROTAREFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim NUMR_TAREFA_ID As Long

Private Sub Form_Load()
   LIMPA_TAREFA
   ABRE_BANCO_MEGASIM NOME_BANCO_DADOS
   SETA_GRID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
      Case vbKeyF10
         GRAVA_TAREFA
         txtCodigo.SetFocus
   End Select
End Sub

Private Sub barTAREFA_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "print"
      Case "limpar"
         LIMPA_TAREFA
         txtCodigo.SetFocus
      Case "del"
         EXCLUI_REGISTRO
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "barTAREFA_ButtonClick"
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then

      NUMR_TAREFA_ID = 0

      If Trim(txtCodigo.Text) = "" Then _
         NUMR_TAREFA_ID = MAX_ID("OSTAREFA_ID", "OSTAREFA", "", "", "", "")
      Else: If Not IsNumeric(txtCodigo.Text) _
               Then NUMR_TAREFA_ID = MAX_ID("OSTAREFA_ID", "OSTAREFA", "", "", "", "")

      txtCodigo.Text = NUMR_TAREFA_ID

      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcodigo_KeyPress"
End Sub

Private Sub txtCodigo_LostFocus()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   MOSTRA_TAREFA

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCodigo_LostFocus"
End Sub

Private Sub TXTDESCRICAO_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTDESCRICAO_KeyPress"
End Sub

Private Sub TXTVALOR_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_KeyPress"
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys ("{tab}")
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPerc_KeyPress"
End Sub

Private Sub TXTDTCAD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_TAREFA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTDTCAD_KeyPress"
End Sub
================================
Private Sub EXCLUI_REGISTRO()
   SQL = "select * from OSTAREFA "
   SQL = SQL & "where OSTAREFA_ID = '" & txtCodigo.Text & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabTemp.EOF Then
      MsgBox "Não Há Registro a ser Excluido.", vbOKOnly, "Atenção !!!"
      Else: TabTemp.Delete
   End If
   TabTemp.Close
   LIMPA_TAREFA
   txtCodigo.SetFocus
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   LISTATAREFA.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSTAREFA "
   SQL = SQL & " order by descricao "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   While Not TabTemp.EOF
      Set Item = LISTATAREFA.ListItems.Add(, "seq." & TabTemp!OSTAREFA_ID, TabTemp!OSTAREFA_ID)

      Item.SubItems(1) = "" & Trim(TabTemp!Descricao)
      Item.SubItems(2) = "" & Format(TabTemp!Valor, "fixed")
      Item.SubItems(3) = "" & Format(TabTemp!perc_comissao, "fixed")

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
=========================
Private Sub LIMPA_TAREFA()
'On Error GoTo ERRO_TRATA

   txtPerc.Text = ""
   txtCodigo.Text = ""
   txtDescricao.Text = ""
   txtValor.Text = ""
   txtDtCad.PromptInclude = False
   txtDtCad.Text = Date
   txtDtCad.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TAREFA"
End Sub

Private Sub MOSTRA_TAREFA()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSTAREFA "
   SQL = SQL & "where OSTAREFA_ID = " & NUMR_TAREFA_ID
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      txtPerc.Text = "" & Format(TabTemp.Fields("perc_comissao").Value, "fixed")
      txtDescricao.Text = "" & Trim(TabTemp.Fields("DESCRICAO").Value)
      txtDtCad.PromptInclude = False
         txtDtCad.Text = "" & TabTemp.Fields("DT_CAD").Value
      txtDtCad.PromptInclude = True
      txtValor.Text = "" & Format(TabTemp.Fields("VALOR").Value, "fixed")
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TAREFA"
End Sub

Private Sub GRAVA_TAREFA()
'On Error GoTo ERRO_TRATA

   If Trim(txtCodigo.Text) = "" Then
      MsgBox "Informe código Tarefa."
      txtCodigo.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao.Text) = "" Then
      MsgBox "Informe descricao Tarefa."
      txtDescricao.SetFocus
      Exit Sub
   End If
   If Trim(txtValor.Text) = "" Then
      MsgBox "Informe valor Tarefa."
      txtValor.SetFocus
      Exit Sub
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSTAREFA "
   SQL = SQL & "where OSTAREFA_ID = " & NUMR_TAREFA_ID
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      SQL = "update OSTAREFA set "
         SQL = SQL & " Descricao = '" & Trim(txtDescricao.Text) & "'"   'Descricao
         SQL = SQL & ", Valor = " & tpMoeda(txtValor.Text)              'Valor
         SQL = SQL & ", PERC_COMISSAO = " & tpMoeda(perc_comissao.Text) 'PERC_COMISSAO
      SQL = SQL & "where OSTAREFA_ID = " & NUMR_TAREFA_ID
      Else
         SQL = "insert into OSTAREFA "
            SQL = SQL & "(OSTAREFA_ID,DESCRICAO,VALOR,PERC_COMISSAO,DT_CAD)"
         SQL = SQL & " values("
            SQL = SQL & NUMR_TAREFA_ID                         'OSTAREFA_ID
            SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'Descricao
            SQL = SQL & "," & tpMoeda(txtValor.Text)           'Valor
            SQL = SQL & "," & tpMoeda(perc_comissao.Text)      'PERC_COMISSAO
            SQL = SQL & ",'" & DMA(txtDtCad.Text) & "'"        'DT_CAD
         SQL = SQL & ")"
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   CONECTA_RETAGUARDA.Execute SQL

   SETA_GRID
   LIMPA_TAREFA

   txtCodigo.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TAREFA"
End Sub
