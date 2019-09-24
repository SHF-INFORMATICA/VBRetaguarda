VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPEDIDOBARRAS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produtos Cadastrados para Código de Barras"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "PEDIDOBARRAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView Lista 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   6324
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
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
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   7232
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Código Barras"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmPEDIDOBARRAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   MOSTRA_RODAPE "Duplo click selecinar item ou tecle entre no item selecionado", "", "", "", ""

   SQL = "select * from PRODUTO "
   SQL = SQL & " where codg_barra = '" & Trim(CRITERIO_A) & "'"
   'SQL = SQL & " and situacao <> 'C' "
   SQL = SQL & " order by situacao"
   SETA_GRID
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   LISTA.ListItems.Clear
   NUMR_SEQ_N = 0
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      NUMR_SEQ_N = NUMR_SEQ_N + 1
      Set item = LISTA.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp.Fields("codg_produto").Value)
      item.SubItems(1) = TabTemp.Fields("descricao").Value
      item.SubItems(2) = TabTemp.Fields("codg_barra").Value

      If Trim(TabTemp.Fields("situacao").Value) = "C" Then
         item.ForeColor = vbRed
         item.ListSubItems(1).ForeColor = vbRed
         item.ListSubItems(2).ForeColor = vbRed
         Else
            item.ForeColor = vbBlack
            item.ListSubItems(1).ForeColor = vbBlack
            item.ListSubItems(2).ForeColor = vbBlack
      End If

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
MsgBox Err.Description
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  OrdenaListView LISTA, ColumnHeader
End Sub

Private Sub Lista_DblClick()
On Error Resume Next

   CRITERIO_A = LISTA.SelectedItem.Text

   Unload Me
End Sub

Private Sub Lista_KeyPress(KeyAscii As Integer)
On Error Resume Next

   If KeyAscii = 13 Then
      KeyAscii = 0
      CRITERIO_A = LISTA.SelectedItem.Text

      Unload Me
   End If
End Sub
