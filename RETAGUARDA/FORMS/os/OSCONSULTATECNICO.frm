VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSConsultaTecnico 
   Caption         =   "Ordem de Serviço Pendente"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OSCONSULTATECNICO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.387
   ScaleMode       =   0  'User
   ScaleWidth      =   21.061
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtValorDig 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbTecnicoAUX 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbTecnico 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   960
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   1693
      ButtonWidth     =   1535
      ButtonHeight    =   1535
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "voltar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3360
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":941B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":AC2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":C007
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":D317
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OSCONSULTATECNICO.frx":E422
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/#### ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4471
      _Version        =   393216
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      DesignWidth     =   11940
      DesignHeight    =   4755
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   15
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   11895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   15
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   11895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tecnico:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
      Height          =   255
      Left            =   8880
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmOSConsultaTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   cmbTecnicoAUX.Clear
   cmbTecnico.Clear

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "select usuario_id, nome from USUARIO "
   SQL = SQL & " where tipo = 9 "   'mecanico
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabDESCR.EOF

      cmbTecnicoAUX.AddItem TabDESCR.Fields("usuario_id").Value
      cmbTecnico.AddItem Trim(TabDESCR.Fields("nome").Value) & "-" & Trim(TabDESCR.Fields("usuario_id").Value)

      TabDESCR.MoveNext
   Wend

   SETA_GRID
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         cmbTecnico.Text = ""
         cmbTecnicoAUX.Text = ""
         txtDtIni.PromptInclude = False
         txtDtFim.PromptInclude = False
         txtDtIni.Text = ""
         txtDtFim.Text = ""
      Case "consultar"
   End Select
End Sub

Private Sub cmbTecnico_Click()
On Error Resume Next

   cmbTecnicoAUX.ListIndex = cmbTecnico.ListIndex
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ERRO_TRATA

    ' Quando clicar uma vez
    ' atribui o valor selecionado
    'AtribuiValorCelula
    OcultarControles

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_Click"
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

   OcultarControles
   ExibirCelula

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

   Select Case KeyAscii
      Case vbKeyReturn  ' Editar ao teclar ENTER
         KeyAscii = 0
         ExibirCelula
      Case vbKeyEscape  ' Cancelar ao pressionar ESC
         KeyAscii = 0
         'AtribuiValorCelula
      Case 32 To 255    ' Editar ao pressinar qualquer tecla
         ExibirCelula
         With txtValorDig
            If .Visible Then
             .Text = Chr$(KeyAscii)
             .SelStart = Len(.Text) + 1
           End If
         End With
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyPress"
End Sub

Private Sub ExibirCelula()
'On Error GoTo ERRO_TRATA

   Static OK As Boolean

   If MSFlexGrid1.Col >= 3 And MSFlexGrid1.Col <= 5 Then
      ' Se for celula fixa , sair
      If MSFlexGrid1.Col <= MSFlexGrid1.FixedCols - 1 Or MSFlexGrid1.Row <= MSFlexGrid1.FixedRows - 1 Then _
         Exit Sub
   
      If OK Then _
         Exit Sub

      OK = True

      OcultarControles

      LastRow = MSFlexGrid1.Row
      LastCol = MSFlexGrid1.Col

      Select Case LastCol
         Case Else
            txtValorDig.Move MSFlexGrid1.CellLeft - Screen.TwipsPerPixelX, MSFlexGrid1.CellTop + MSFlexGrid1.Top - Screen.TwipsPerPixelY, MSFlexGrid1.CellWidth + Screen.TwipsPerPixelX * 2, MSFlexGrid1.CellHeight + Screen.TwipsPerPixelY * 2
            txtValorDig.Text = MSFlexGrid1.Text

            If Len(MSFlexGrid1.Text) = 0 Then _
               If LastRow > 1 Then _
                  txtValorDig.Text = MSFlexGrid1.TextMatrix(LastRow - 1, LastCol)

            txtValorDig.Visible = True

            If txtValorDig.Visible Then
               txtValorDig.ZOrder
               txtValorDig.SetFocus
            End If
      End Select
   
      ControlVisible = True

      OK = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ExibirCelula"
End Sub

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorDig.SelStart = 0
   txtValorDig.SelLength = Len(txtValorDig)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_GotFocus"
End Sub

Private Sub txtValorDig_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         OcultarControles
         MSFlexGrid1.SetFocus
      Case vbKeyUp
         OcultarControles
         'move para a cima celula.
         With MSFlexGrid1
            If .Row > 1 Then
                .Row = .Row - 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
      Case vbKeyDown
         OcultarControles
         With MSFlexGrid1
             If .Row + 1 < .Rows Then
                .Row = .Row + 1
                '.Col = 0
               Else
                .Row = 1
                '.Col = 0
            End If
         End With

         ExibirCelula
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyDown"
End Sub

Private Sub txtValorDig_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   ' ao pressionar ENTER aceitar a entrada de dados
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      If LastCol = 5 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      CFOP_ID_N = 0 & txtValorDig.Text

      Dim CFOP_ID_ANT As Double

      CFOP_ID_ANT = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
      'PEDIDO_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
      PRODUTO_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)
      SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)

      'AtribuiValorCelula
      'ProximaCelula
      OcultarControles

      If CFOP_ID_N > 0 Then
         SQL = "update PEDIDOITEM set "
         SQL = SQL & " cfop_id = " & CFOP_ID_N

         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         CFOP_ID_ANT = 0
         PRODUTO_ID_N = 0
         SEQ_ID_N = 0

         SETA_GRID
      End If

      With MSFlexGrid1
         If .Row + 1 < .Rows Then
            .Row = .Row + 1
            '.Col = 0
            Else
               .Row = 1
               '.Col = 0
         End If
      End With
      txtValorDig.Text = ""
      NaturezaOperacao_A = ""
      MSFlexGrid1.SetFocus
      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
            Else
               If KeyAscii = 8 Or KeyAscii = 44 Then
                  Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_KeyPress"
End Sub

Private Sub ProximaCelula()
'On Error GoTo ERRO_TRATA

   If MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then
      MSFlexGrid1.Col = MSFlexGrid1.Col + 1
      Else
         MSFlexGrid1.Col = 1
         If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
             MSFlexGrid1.Row = MSFlexGrid1.Row + 1
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ProximaCelula"
End Sub

Private Sub AtribuiValorCelula()
'On Error GoTo ERRO_TRATA

   Dim texto As String

   ' atribuir o texto anterior a celula
   Select Case LastCol
      Case 5
         texto = txtValorDig.Text
         MSFlexGrid1.TextMatrix(LastRow, LastCol) = Trim(texto)
         MSFlexGrid1.CellForeColor = vbRed
         MSFlexGrid1.CellFontBold = True
         MSFlexGrid1.CellBackColor = &H8000000F
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "AtribuiValorCelula"
End Sub

Private Sub OcultarControles()
'On Error GoTo ERRO_TRATA

   'Ocultar o controle textbox
   txtValorDig.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "OcultarControles"
End Sub
'===================================
Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   Dim Coluna, Linha, Largura_Campo

   MSFlexGrid1.Clear
   MSFlexGrid1.Visible = False
   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select OS.OS_ID, OS.DT_OS, OS.TIPO_OS, OS.SITUACAO_OS, OS.CLIENTE, "
   SQL = SQL & " OSSERVICO.OSSERVICO_ID, OSSERVICO.OSTAREFA_ID, OSSERVICO.DT_CAD, "
   SQL = SQL & " OSSERVICO.SITUACAO, OSSERVICO.RESPONSAVEL_ID, OSSERVICO.VALOR_SERVICO, "
   SQL = SQL & " OSSERVICO.DESCRICAO, OSSERVICO.DESCONTO_SERVICO, OSSERVICO.DT_FIM , "
   SQL = SQL & " OSSERVICO.DT_INICIO, USUARIO.NOME, USUARIO.CPF"
   SQL = SQL & " from OS WITH (NOLOCK)"
   SQL = SQL & " Inner Join OSSERVICO WITH (NOLOCK)"
   SQL = SQL & " ON OS.OS_ID = OSSERVICO.OS_ID "
   SQL = SQL & " INNER JOIN USUARIO WITH (NOLOCK)"
   SQL = SQL & " ON OSSERVICO.RESPONSAVEL_ID = USUARIO.USUARIO_ID"

SQL = SQL & " where situacao in ('P','E','O')"

   If Trim(cmbTecnicoAUX.Text) <> "" Then _
      If IsNumeric(cmbTecnicoAUX.Text) Then _
         SQL = SQL & " and RESPONSAVEL_ID = " & cmbTecnicoAUX.Text

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabTemp.Fields.Count

      ReDim largura_coluna(0 To TabTemp.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabTemp.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabTemp.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabTemp.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabTemp.EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabTemp.Fields.Count - 1
            If Coluna = 2 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabTemp.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  If Coluna = 3 Or Coluna = 4 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabTemp.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabTemp.Fields(Coluna).Value)
                  End If
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabTemp.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabTemp.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

         TabTemp.MoveNext
         Linha = Linha + 1
      Loop

      'define a largura das colunas do grid
      For Coluna = 0 To MSFlexGrid1.Cols - 1
         MSFlexGrid1.ColWidth(Coluna) = largura_coluna(Coluna) + 240
      Next Coluna

      MSFlexGrid1.ColWidth(0) = 0
      MSFlexGrid1.Refresh

      MSFlexGrid1.BackColor = vbWhite
      MSFlexGrid1.ForeColor = vbBlue

'CellFontName        - Define o nome da fonte para uma célula
'CellFontSize        - Define o tamanho da fonte para a célula
'CellFontBold        - Define se a fonte aparece em negrito.
'CellFontItalic      - Define se a fonte aparece em itálico.
'CellFontUnderline   - Define se a fonte aparece sublinhada.

'OS.OS_ID
      MSFlexGrid1.ColWidth(0) = 1000
      MSFlexGrid1.ColAlignment(0) = 0

'OS.DT_OS
      MSFlexGrid1.ColWidth(1) = 2000
      MSFlexGrid1.ColAlignment(1) = 0

'OS.TIPO_OS
      MSFlexGrid1.ColWidth(2) = 500
      MSFlexGrid1.ColAlignment(2) = 7

'OS.SITUACAO_OS
      MSFlexGrid1.ColWidth(3) = 500
      MSFlexGrid1.ColAlignment(3) = 7

'OS.CLIENTE
      MSFlexGrid1.ColWidth(4) = 5000
      MSFlexGrid1.ColAlignment(4) = 7

'OSSERVICO.OSSERVICO_ID
      MSFlexGrid1.ColWidth(5) = 1
      MSFlexGrid1.ColAlignment(5) = 7

'OSSERVICO.OSTAREFA_ID
      MSFlexGrid1.ColWidth(6) = 1
      MSFlexGrid1.ColAlignment(6) = 0

'OSSERVICO.DT_CAD
      MSFlexGrid1.ColWidth(7) = 2000
      MSFlexGrid1.ColAlignment(7) = 0

'OSSERVICO.SITUACAO
      MSFlexGrid1.ColWidth(8) = 500
      MSFlexGrid1.ColAlignment(8) = 0

'OSSERVICO.RESPONSAVEL_ID
      MSFlexGrid1.ColWidth(9) = 500
      MSFlexGrid1.ColAlignment(9) = 0

'OSSERVICO.VALOR_SERVICO
      MSFlexGrid1.ColWidth(10) = 2000
      MSFlexGrid1.ColAlignment(10) = 0

'OSSERVICO.DESCRICAO
      MSFlexGrid1.ColWidth(11) = 6000
      MSFlexGrid1.ColAlignment(11) = 0

'OSSERVICO.DESCONTO_SERVICO
      MSFlexGrid1.ColWidth(12) = 1000
      MSFlexGrid1.ColAlignment(12) = 0

'OSSERVICO.DT_INICIO
     MSFlexGrid1.ColWidth(13) = 2000
     MSFlexGrid1.ColAlignment(13) = 0

'OSSERVICO.DT_FIM
      MSFlexGrid1.ColWidth(14) = 2000
      MSFlexGrid1.ColAlignment(14) = 0

'USUARIO.NOME
      MSFlexGrid1.ColWidth(15) = 3000
      MSFlexGrid1.ColAlignment(15) = 0

'USUARIO.CPF
      MSFlexGrid1.ColWidth(16) = 1
      MSFlexGrid1.ColAlignment(16) = 0
   
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   MSFlexGrid1.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
