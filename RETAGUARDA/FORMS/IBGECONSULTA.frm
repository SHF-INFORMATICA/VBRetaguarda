VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmIBGECONSULTA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta IBGE"
   ClientHeight    =   5985
   ClientLeft      =   3525
   ClientTop       =   2685
   ClientWidth     =   8565
   Icon            =   "IBGECONSULTA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8565
   Begin VB.TextBox txtibge 
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
      Left            =   80
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
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
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
   End
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":739E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":76BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IBGECONSULTA.frx":80D0
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
      Width           =   8565
      _ExtentX        =   15108
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
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3000
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
               Picture         =   "IBGECONSULTA.frx":8AE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IBGECONSULTA.frx":9F0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IBGECONSULTA.frx":AF99
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LISTA 
      Height          =   4365
      Left            =   60
      TabIndex        =   4
      Top             =   1560
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   7699
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   5200
      Left            =   0
      Top             =   750
      Width           =   8535
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1830
      TabIndex        =   7
      Top             =   840
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7680
      TabIndex        =   6
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IBGE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   75
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmIBGECONSULTA"
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
      Case vbKeyF9: LIMPA_IBGE
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

Private Sub Lista_DblClick()
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   If LISTA.SelectedItem.Text <> "" Then
      CRITERIO_A = LISTA.SelectedItem.Text
      SQL3 = "" & LISTA.SelectedItem.ListSubItems.item(1).Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LISTA_DblClick"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         LIMPA_IBGE
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtCidade_LostFocus()
   txtCidade.BackColor = &HFFFFFF
End Sub

Private Sub txtIBGE_Change()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Chr$(39) & txtIBGE.Text & Chr(39)
   SQL = "select * from IBGE WHERE IBGE_ID like " & CRITERIO_A
   SQL = SQL & " and IBGE_ID is not null "
   SQL = SQL & " and municipio is not null "
   SQL = SQL & " and estado is not null "
   CRITERIO_A = txtIBGE.Text
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIBGE_Change"
End Sub

Private Sub txtibge_GotFocus()
'On Error GoTo ERRO_TRATA

   txtIBGE.SelStart = 0
   txtIBGE.SelLength = Len(txtIBGE)
   txtIBGE.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtIBGE_GotFocus"
End Sub

Private Sub txtCidade_Change()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Chr$(39) & txtCidade.Text & "%" & Chr(39)
   SQL = "select * from IBGE WHERE municipio like " & CRITERIO_A
   SQL = SQL & " and IBGE_ID is not null "
   SQL = SQL & " and municipio is not null "
   SQL = SQL & " and estado is not null "
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCidade_Change"
End Sub

Private Sub txtcidade_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCidade.SelStart = 0
   txtCidade.SelLength = Len(txtCidade)
   txtCidade.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCidade_GotFocus"
End Sub

Private Sub txtIbge_LostFocus()
   txtIBGE.BackColor = &HFFFFFF
End Sub

Private Sub txtUF_Change()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = Chr$(39) & txtUF.Text & "%" & Chr(39)
   SQL = "select * from IBGE WHERE estado like " & CRITERIO_A
   SQL = SQL & " and IBGE_ID is not null "
   SQL = SQL & " and municipio is not null "
   SQL = SQL & " and estado is not null "
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_Change"
End Sub

Private Sub txtuf_GotFocus()
'On Error GoTo ERRO_TRATA

   txtUF.SelStart = 0
   txtUF.SelLength = Len(txtUF)
   txtUF.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_GotFocus"
End Sub

Private Sub txtUF_LostFocus()
   txtUF.BackColor = &HFFFFFF
End Sub

Private Sub LIMPA_IBGE()
'On Error GoTo ERRO_TRATA

   txtCidade.Text = ""
   txtUF.Text = ""
   txtIBGE.Text = ""
   LISTA.ListItems.Clear
   txtIBGE.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_IBGE"
End Sub
   
Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   NUMR_CONSULTA_N = 0
   LISTA.ListItems.Clear

   If TabTemp.State = 1 Then TabTemp.Close
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      Set item = LISTA.ListItems.Add(, "seq." & NUMR_CONSULTA_N, TabTemp!IBGE_ID)
      item.SubItems(1) = TabTemp!Municipio
      item.SubItems(2) = TabTemp!Estado
      TabTemp.MoveNext
   Wend
   TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
