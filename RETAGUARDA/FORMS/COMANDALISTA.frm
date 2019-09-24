VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCOMANDALISTA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comandas Pendentes"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   Icon            =   "COMANDALISTA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   9150
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstCMD 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   5900
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Comanda"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Atendente"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "LISTAGEM COMANDAS PENDENTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmCOMANDALISTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   lstCMD.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PEDIDOTEMP.PEDIDO_ID, PEDIDOTEMP.ESTABELECIMENTO_ID, PEDIDOTEMP.CARTAOBARRA_ID, PEDIDOTEMP.USUARIO_ID, PEDIDOTEMP.DT_PEDIDO, "
   SQL = SQL & " CARTAOBARRA.CODIGO_BARRA, CARTAOBARRA.DESCRICAO, CARTAOBARRA.DTCAD, CARTAOBARRA.Status"
   SQL = SQL & " from PEDIDOTEMP "
   SQL = SQL & " LEFT OUTER JOIN CARTAOBARRA "
   SQL = SQL & " ON PEDIDOTEMP.CARTAOBARRA_ID = CARTAOBARRA.CARTAOBARRA_ID"
   SQL = SQL & " where dt_pedido <> '" & Trim(Date) & "'"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstCMD.ListItems.Add(, "seq." & Trim(TabTemp.Fields("cartaobarra_id").Value), Trim(TabTemp.Fields("cartaobarra_id").Value))
      item.SubItems(1) = "" & DMA(TabTemp.Fields("dt_pedido").Value)
      item.SubItems(2) = "" & Trim(TabTemp.Fields("descricao").Value)
      item.SubItems(3) = "" & TRAZ_NOME_USUARIO(Trim(TabTemp.Fields("usuario_id").Value))
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         Unload frmCOMANDALISTA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KEYDOWN"
End Sub
