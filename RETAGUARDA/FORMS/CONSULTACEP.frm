VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCONSULTACEP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Cep"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   915
   ClientWidth     =   8640
   Icon            =   "CONSULTACEP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUF 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtCidade 
      DataSource      =   "DataCep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
   End
   Begin MSMask.MaskEdBox txtCep 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####-###"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6000
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTACEP.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTACEP.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CONSULTACEP.frx":80C9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   4545
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   8017
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cep"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cidade"
         Object.Width           =   9172
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "UF"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cep:"
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
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF:"
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
      Height          =   225
      Left            =   7680
      TabIndex        =   4
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
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
      Height          =   225
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmCONSULTACEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9: LIMPA_CEP
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
      Case "limpar"
         LIMPA_CEP
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Lista_DblClick()
'On Error GoTo ERRO_TRATA

   If LISTA.SelectedItem.Text <> "" Then
      CRITERIO_A = LISTA.SelectedItem.Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LISTA_DblClick"
End Sub

Private Sub txtCep_Change()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Chr$(39) & txtCep.Text & "%" & Chr(39)
   SQL = "select * from cep where cep_ID like " & CRITERIO_A
   SQL = SQL & " and cep_ID is not null "
   SQL = SQL & " and cidade is not null "
   SQL = SQL & " and uf is not null "
   CRITERIO_A = txtCep.Text
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_Change"
End Sub

Private Sub txtCep_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Digite o Cep"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "Para selecionar Duplo Clik no Grid"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_GotFocus"
End Sub

Private Sub txtCidade_Change()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Chr$(39) & txtCidade.Text & "%" & Chr(39)
   SQL = "select * from cep WHERE cidade like " & CRITERIO_A
   SQL = SQL & " and cep_ID is not null "
   SQL = SQL & " and cidade is not null "
   SQL = SQL & " and uf is not null "
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCidade_Change"
End Sub

Private Sub txtcidade_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Digite a Cidade"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "Para selecionar Duplo Clik no Grid"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCidade_GotFocus"
End Sub

Private Sub txtUF_Change()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Chr$(39) & txtUF.Text & "%" & Chr(39)
   SQL = "select * from cep WHERE uf like " & CRITERIO_A
   SQL = SQL & " and cep_ID is not null "
   SQL = SQL & " and cidade is not null "
   SQL = SQL & " and uf is not null "
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_Change"
End Sub

Private Sub txtuf_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Digite o Estado"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "Para selecionar Duplo Clik no Grid"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_GotFocus"
End Sub

Private Sub LIMPA_CEP()
'On Error GoTo ERRO_TRATA

   txtCidade.Text = ""
   txtUF.Text = ""
   txtCep.Text = ""
   LISTA.ListItems.Clear
   txtCep.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_CEP"
End Sub
   
Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   NUMR_CONSULTA_N = 0
   LISTA.ListItems.Clear
   If TabTemp.State = 1 Then TabTemp.Close
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      Set item = LISTA.ListItems.Add(, "seq." & NUMR_CONSULTA_N, TabTemp!CEP_ID)
      item.SubItems(1) = TabTemp!CIDADE
      item.SubItems(2) = TabTemp!UF
      TabTemp.MoveNext
   Wend
   TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
