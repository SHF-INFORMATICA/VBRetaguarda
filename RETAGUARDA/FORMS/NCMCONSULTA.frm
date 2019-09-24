VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNCMConsulta 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Consuta NCM"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NCMCONSULTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8280
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCodgNCM 
      Height          =   360
      Left            =   6960
      MaxLength       =   8
      TabIndex        =   1
      ToolTipText     =   "Informe Codigo NCM do produto"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtDesc 
      Height          =   360
      Left            =   1980
      MaxLength       =   100
      TabIndex        =   0
      ToolTipText     =   "Informe "
      Top             =   960
      Width           =   4215
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   1270
      ButtonWidth     =   2593
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   7
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3480
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":9331
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":AA2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":B9E3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":CAEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":DC88
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":EEBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NCMCONSULTA.frx":F30C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstNCM 
      Height          =   1995
      Left            =   45
      TabIndex        =   2
      Top             =   4440
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   3519
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Produto"
         Object.Width           =   17639
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "NCM"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView lstProduto 
      Height          =   2000
      Left            =   45
      TabIndex        =   6
      Top             =   1920
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   3519
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Produto"
         Object.Width           =   17639
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "NCM"
         Object.Width           =   5292
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   8280
      DesignHeight    =   6570
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela NCM Sugerida"
      Height          =   240
      Left            =   75
      TabIndex        =   8
      Top             =   4080
      Width           =   8205
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos Cadastrados"
      Height          =   240
      Left            =   3105
      TabIndex        =   7
      Top             =   1560
      Width           =   2070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   10320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   10320
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   10320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NCM:"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6360
      TabIndex        =   4
      Top             =   975
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição Produto:"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1800
   End
End
Attribute VB_Name = "frmNCMConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "limpar"
         txtCodgNCM.Text = ""
         txtDesc.Text = ""
         lstNCM.ListItems.Clear
         lstProduto.ListItems.Clear
         txtDesc.SetFocus
      Case "voltar"
         CRITERIO_A = ""
         Unload Me
      Case "consultar"
         If Trim(txtDesc.Text) <> "" Or Trim(txtCodgNCM.Text) <> "" Then
            CONSULTA_PRODUTO
            CONSULTA_NCM
         End If
      Case "gravar"
         GRAVA_NCM_PRODUTO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub lstNCM_Click()
'On Error GoTo ERRO_TRATA

   txtCodgNCM.Text = "" & lstNCM.SelectedItem.ListSubItems.item(1).Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstNCM_Click"
End Sub

Private Sub LSTNCM_DblClick()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""
   If Not IsNull(lstNCM.SelectedItem.Text) Then
      If Trim(lstNCM.SelectedItem.ListSubItems.item(1).Text) = "" Then
         PEDIDO_ID_N = 0
         Exit Sub
      End If
      CRITERIO_A = lstNCM.SelectedItem.ListSubItems.item(1).Text
      Unload Me
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "lstNCM_DblClick"
End Sub

Private Sub txtDesc_GotFocus()
   txtDesc.SelStart = 0
   txtDesc.SelLength = Len(txtDesc)
   txtDesc.BackColor = &HC0FFFF
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtDesc.Text) <> "" Then
         CONSULTA_PRODUTO
         CONSULTA_NCM
      End If
   End If
End Sub

Private Sub txtDesc_LostFocus()
   txtDesc.BackColor = &HFFFFFF
End Sub

Private Sub txtCodgNCM_GotFocus()
   txtCodgNCM.SelStart = 0
   txtCodgNCM.SelLength = Len(txtCodgNCM)
   txtCodgNCM.BackColor = &HC0FFFF
End Sub

Private Sub txtCodgNCM_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtCodgNCM.Text) <> "" Then
         CONSULTA_PRODUTO
         CONSULTA_NCM
      End If
   End If

End Sub

Private Sub txtCodgNCM_LostFocus()
   txtCodgNCM.BackColor = &HFFFFFF
End Sub
'===========
Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   'If Trim(txtDesc.Text) = "" Then _
      Exit Sub

   SQL = "select produto_id,codg_produto,descricao,familiaproduto_id,codg_ncm from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where produto_id is not null"

   If Trim(txtCodgNCM.Text) = "" Then
      SQL = SQL & " and descricao like '" & Trim(txtDesc.Text) & "%" & "'"
      Else: SQL = SQL & " and codg_ncm like '" & Trim(txtCodgNCM.Text) & "%" & "'"
   End If

   SQL = SQL & " order by descricao"
   
   SETA_GRID_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_NCM"
End Sub

Sub SETA_GRID_PRODUTO()
'On Error GoTo ERRO_TRATA

   lstProduto.Visible = False
   lstProduto.ListItems.Clear

   If TabProduto.State = 1 Then _
      TabProduto.Close

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   Me.Enabled = True
   While Not TabProduto.EOF
      DoEvents

      Set item = lstProduto.ListItems.Add(, "seq." & TabProduto.Fields("produto_id").Value, Trim(TabProduto.Fields("codg_produto").Value))
      
      item.SubItems(1) = "" & Trim(TabProduto!DESCRICAO)
      item.SubItems(2) = "" & Trim(TabProduto.Fields("codg_ncm").Value)

      If item.SubItems(2) = "" Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed
         item.ListSubItems(2).ForeColor = vbRed
      End If
      If item.SubItems(2) <> "" Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbBlue
         item.ListSubItems(2).ForeColor = vbBlue
      End If

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close
   SQL = ""
   lstProduto.Visible = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_PRODUTO"
End Sub

Sub CONSULTA_NCM()
'On Error GoTo ERRO_TRATA

   'If Trim(txtDesc.Text) = "" Then _
      Exit Sub

   SQL = "select * from TABNCM WITH (NOLOCK)"
   SQL = SQL & " where codg_ncm is not null"

   If Trim(txtCodgNCM.Text) = "" Then
      SQL = SQL & " and descricao like '" & Trim(txtDesc.Text) & "%" & "'"
      Else: SQL = SQL & " and codg_ncm like '" & Trim(txtCodgNCM.Text) & "%" & "'"
   End If

   SQL = SQL & " order by descricao"
   
   SETA_GRID_NCM

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_NCM"
End Sub

Sub SETA_GRID_NCM()
'On Error GoTo ERRO_TRATA

   lstNCM.Visible = False
   lstNCM.ListItems.Clear
   CONT_N = 0

   If TabProduto.State = 1 Then _
      TabProduto.Close

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   Me.Enabled = True
   While Not TabProduto.EOF
      DoEvents
      CONT_N = 1 + CONT_N

      Set item = lstNCM.ListItems.Add(, "seq." & CONT_N, Trim(TabProduto.Fields("DESCRICAO").Value))
      
      item.SubItems(1) = "" & Trim(TabProduto.Fields("codg_ncm").Value)

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

   lstNCM.Visible = True

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_NCM"
End Sub

Sub GRAVA_NCM_PRODUTO()
'On Error GoTo ERRO_TRATA

   If Trim(lstNCM.SelectedItem.ListSubItems.item(1).Text) = "" Then _
      Exit Sub

   Dim IntInicio  As Integer
   Dim strSQL     As String
   Dim i          As Integer

   If lstProduto.ListItems.Count > 0 Then
      For i = lstProduto.ListItems.Count To 1 Step -1
         If lstProduto.ListItems(i).Checked = True Then
            If Trim(lstProduto.ListItems(i).Text) <> "" Then
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               SQL = "select produto_id from PRODUTO "
               SQL = SQL & " where codg_produto = '" & Trim(lstProduto.ListItems(i).Text) & "'"
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabConsulta.EOF Then
                  If TabConsulta.State = 1 Then _
                     TabConsulta.Close

                  SQL = "update produto set "
                  SQL = SQL & " codg_ncm = '" & Trim(lstNCM.SelectedItem.ListSubItems.item(1).Text) & "'"
                  SQL = SQL & " where codg_produto = '" & Trim(lstProduto.ListItems(i).Text) & "'"
                  CONECTA_RETAGUARDA.Execute SQL
               End If
               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               DoEvents
            End If
         End If
      Next i
   End If
   CONSULTA_PRODUTO

Exit Sub
ERRO_TRATA:
   Me.Enabled = True
   Me.KeyPreview = True
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_NCM_PRODUTO"
End Sub
