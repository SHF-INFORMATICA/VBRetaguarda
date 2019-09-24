VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSServicoConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Serviço"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "consultatarefa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   $"consultatarefa.frx":5C12
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   9255
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         MaxLength       =   8
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   5775
      End
   End
   Begin MSComctlLib.ListView LISTATAREFA 
      Height          =   4785
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8440
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
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
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
         Text            =   "Valor Tarefa"
         Object.Width           =   2382
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "desconto"
         Text            =   "Comissão"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "dtemis"
         Text            =   "Obs."
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Key             =   "status"
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
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
            Picture         =   "consultatarefa.frx":5C9F
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultatarefa.frx":60F3
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultatarefa.frx":640F
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultatarefa.frx":6863
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultatarefa.frx":6CB7
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultatarefa.frx":6FD7
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "consultatarefa.frx":742B
            Key             =   "IMG7"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barTAREFA 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmOSSERVICOCONSULTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS

   SETA_GRID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         LIMPA_TAREFA
         txtCodigo.SetFocus
   End Select
End Sub

Private Sub barTAREFA_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "print"
      Case "limpar"
         LIMPA_TAREFA
         txtCodigo.SetFocus
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub Listatarefa_DblClick()
   SQL3 = ""
   If Not IsNull(LISTATAREFA.SelectedItem.Text) Then
      SQL3 = LISTATAREFA.SelectedItem.Text
      Unload Me
   End If
End Sub

Private Sub txtDesc_Change()
   CRITERIO_A = Chr$(39) & txtDesc.Text & "*" & Chr(39)
   SETA_GRID
End Sub

Private Sub txtcodigo_Change()
   'CRITERIO_A = Chr$(39) & txtCodigo.Text & "*" & Chr(39)
   SETA_GRID
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(LISTATAREFA.SelectedItem.Text) Then
         KeyAscii = 0
         CODG_PROD_A = LISTATAREFA.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(LISTATAREFA.SelectedItem.Text) Then
         KeyAscii = 0
         CODG_PROD_A = LISTATAREFA.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(LISTATAREFA.SelectedItem.Text) Then
         KeyAscii = 0
         CODG_PROD_A = LISTATAREFA.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtvalor_Change()
   'CRITERIO_A = Chr$(39) & txtvalor.Text & "*" & Chr(39)
   SETA_GRID
End Sub

Private Sub LIMPA_TAREFA()
   txtCodigo.Text = ""
   txtDesc.Text = ""
   txtValor.Text = ""
   txtDesc.SetFocus
End Sub

Private Sub SETA_GRID()
   LISTATAREFA.ListItems.Clear
   NUMR_CONSULTA_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from OSTAREFA "
   SQL = SQL & " where descricao <> '' "

   If Trim(txtDesc.Text) <> "" Then _
      SQL = SQL & " and descricao like " & Trim(txtDesc.Text) & "%"

   If Trim(txtCodigo.Text) <> "" Then _
      If IsNumeric(txtCodigo.Text) Then _
         SQL = SQL & " and OSTAREFA_ID = " & txtCodigo.Text

   If txtValor.Text <> "" Then _
      SQL = SQL & " and valor_tarefa >= " & txtValor.Text

   SQL = SQL & " order by descricao"

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = LISTATAREFA.ListItems.Add(, "seq." & TabTemp.Fields("ostarefa_id").Value, TabTemp.Fields("ostarefa_id").Value)
      item.SubItems(1) = TabTemp.Fields("descricao").Value
      item.SubItems(2) = Format(TabTemp.Fields("valor").Value, strFormatacao2Digitos)
      item.SubItems(3) = Format(TabTemp.Fields("perc_comissao").Value, strFormatacao2Digitos)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
End Sub
