VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmDisplayPromocao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produtos Promoção"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DisplayPromocao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11400
      DesignHeight    =   6495
   End
   Begin MSComctlLib.ListView lstProgramacao 
      Height          =   5895
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   10398
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Produto"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DtIni"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DtFim"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Duração"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "MSG Display"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "MSG Promocional"
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "frmDisplayPromocao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   SETA_GRID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF6
         If Trim(lstPedidos.SelectedItem.ListSubItems.item(9).Text) = "DC" Then 'Devolução de Entrada
         End If
      Case vbKeyEscape
         Unload Me
   End Select

BlockInput False  'Desbloqueia o teclado
Exit Sub
ERRO_TRATA:
   BlockInput False  'Desbloqueia o teclado
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   lstProgramacao.ListItems.Clear

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PRODUTOPROMOCAO.*, "
   SQL = SQL & " PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
   SQL = SQL & " PRODUTO.UNIDADE_MEDIDA, Produto.Tipo_Prod"
   SQL = SQL & " from PRODUTO WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PRODUTOPROMOCAO WITH (NOLOCK) "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PRODUTOPROMOCAO.PRODUTO_ID"

   SQL = SQL & " where PRODUTOPROMOCAO_ID > 0 "

   'SQL = SQL = " and dataini >= '" & DMA(Date) & "'"

   If PRODUTO_ID_N > 0 Then _
      SQL = SQL & " and PRODUTOPROMOCAO.produto_id = " & PRODUTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Set item = lstProgramacao.ListItems.Add(, "seq." & TabTemp.Fields("PRODUTOPROMOCAO_id").Value, TabTemp.Fields("PRODUTOPROMOCAO_id").Value)
      item.SubItems(1) = "" & Trim(TabTemp.Fields("codg_produto").Value)
      item.SubItems(2) = "" & Trim(TabTemp.Fields("descricao").Value)
      item.SubItems(3) = "" & Trim(TabTemp.Fields("DATAINI").Value)
      item.SubItems(4) = "" & Trim(TabTemp.Fields("dAtAfim").Value)
      item.SubItems(5) = ""
      If Not IsNull(TabTemp.Fields("dAtAfim").Value) And Not IsNull(TabTemp.Fields("dAtAini").Value) Then _
         item.SubItems(5) = "" & CONVERTE_TEMPO(TabTemp.Fields("dAtAini").Value, TabTemp.Fields("dAtAfim").Value)
      item.SubItems(6) = "" & Trim(TabTemp.Fields("msgdisplay").Value)
      item.SubItems(7) = "" & Trim(TabTemp.Fields("msgpromocao").Value)

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Function CONVERTE_TEMPO(DATA_INI_D As String, DATA_FIM_D As String) As String
   Dim d1, d2, d3, Segundo_n, d5  As Single
   Dim Minuto_N            As Long
   Dim Hora_N              As Long
   Dim Dia_N               As Long
   Dim Data_ini_Vaca       As Date
   Dim Data_fim_Vaca       As Date

   If IsDate(DATA_INI_D) Then
      Data_ini_Vaca = DATA_INI_D
      Else: Data_ini_Vaca = Format(DATA_INI_D, "##/##/#### ##:##:##")
   End If
   If IsDate(DATA_FIM_D) Then
      Data_fim_Vaca = DATA_FIM_D
      Else: Data_fim_Vaca = Format(DATA_FIM_D, "##/##/#### ##:##:##")
   End If

   d1 = DateDiff("d", Data_ini_Vaca, Data_fim_Vaca)
   d2 = DateDiff("m", Data_ini_Vaca, Data_fim_Vaca)
   d3 = DateDiff("yyyy", Data_ini_Vaca, Data_fim_Vaca)
   Segundo_n = DateDiff("s", Data_ini_Vaca, Data_fim_Vaca)

   Msg = " Sua idade e : " & vbCrLf
   Msg = Msg & " ============================== " & vbCrLf
   Msg = Msg & " Em dias : " & d1 & " dias " & vbCrLf
   Msg = Msg & " Em meses : " & d2 & " meses " & vbCrLf
   Msg = Msg & " Em anos : " & d3 & " anos " & vbCrLf
   Msg = Msg & " Em segundos : " & Segundo_n & " segundos " & vbCrLf

   CONVERTE_TEMPO = ""

'converção para minutos
   Minuto_N = 0
   While Segundo_n >= 60
      Segundo_n = Segundo_n - 60
      Minuto_N = Minuto_N + 1
   Wend
   
'converção para horas
   Hora_N = 0
   While Minuto_N >= 60
      Minuto_N = Minuto_N - 60
      Hora_N = Hora_N + 1
   Wend

'converção para dia
   Dia_N = 0
   While Hora_N >= 60
      Hora_N = Hora_N - 60
      Dia_N = Dia_N + 1
   Wend

   'txtDuracao.Text = Minuto_N & ":" & Segundo_n
   CONVERTE_TEMPO = "" & Minuto_N & ":" & Segundo_n
   
'MsgBox "Dia = " & Dia_n & " ; Horas = " & Hora_n & " ; Minutos = " & Minuto_n & " ; Segundos = " & Segundo_n

End Function
