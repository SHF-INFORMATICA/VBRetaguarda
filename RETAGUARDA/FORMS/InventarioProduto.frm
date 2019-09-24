VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmInventarioProduto 
   Caption         =   "Relatório Inventário/Produto/Contador"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InventarioProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   12465
   Begin VB.CheckBox chkInativos 
      Caption         =   "Inativos"
      Height          =   240
      Left            =   1920
      TabIndex        =   15
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox chkZerados 
      Caption         =   "Qtde Zerados"
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.ComboBox cmbTabela 
      Height          =   360
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "Selecione o grupo do produto."
      Top             =   1320
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ComboBox cmbFamiliaAux 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   2130
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Selecione o grupo do produto."
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdConsulta 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   2085
      Picture         =   "InventarioProduto.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   405
   End
   Begin VB.TextBox txtRazao 
      DataField       =   "Nome"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2535
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   2249
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
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "entrada"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   6960
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   5040
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioProduto.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioProduto.frx":77AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioProduto.frx":883D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioProduto.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InventarioProduto.frx":AA12
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   405
      Left            =   135
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView lstCampos 
      Height          =   1335
      Left            =   5880
      TabIndex        =   9
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2355
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      TextBackground  =   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Campos"
         Object.Width           =   1764
      EndProperty
   End
   Begin Threed.SSCommand cmdGrid 
      Height          =   495
      Left            =   9600
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Mostrar Grid"
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdExportar 
      Height          =   495
      Left            =   11040
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Exportar"
      PictureAlignment=   6
   End
   Begin MSDataGridLib.DataGrid grdCampos 
      Bindings        =   "InventarioProduto.frx":BB1D
      Height          =   4575
      Left            =   45
      TabIndex        =   12
      Top             =   2880
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      WrapCellPointer =   -1  'True
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoCampos 
      Height          =   330
      Left            =   9840
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Grid Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblgrupo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Família Produto:"
      Height          =   240
      Left            =   435
      TabIndex        =   8
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1155
   End
End
Attribute VB_Name = "frmInventarioProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim SQL_CONSULTA As String

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   CARREGA_TABELAS
   LIMPA_TUDO
   PreencheComboGrupo cmbFamilia

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "entrada"
         IMPRIMIR_REL
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TUDO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAUX.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   TIPO_PESSOA_CADASTRO = "FORNECEDOR"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = CNPJCPF_A
      txtCNPJCPF.PromptInclude = True
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsulta_Click"
End Sub


Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False

   If Trim(txtCNPJCPF.Text = "") Then _
      If txtCNPJCPF.Mask = "" Then _
         txtCNPJCPF.Mask = "##############"

   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         TIPO_PESSOA_CADASTRO = "FORNECEDOR"
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
            txtCNPJCPF.PromptInclude = True
         End If
         txtCNPJCPF.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
        txtCNPJCPF.PromptInclude = False
        If txtCNPJCPF.Text = "" Then
           Else
              If Len(txtCNPJCPF.Text) > 0 Then
                 Select Case Len(txtCNPJCPF.Text)
                    Case Is = 11
                      If Not CALCULACPF(txtCNPJCPF.Text) Then
                         MsgBox "CPF com DV incorreto !!!"
                         txtCNPJCPF.PromptInclude = False
                         txtCNPJCPF = ""
                         txtCNPJCPF.SetFocus
                         Exit Sub
                      End If
                    Case Is = 14
                      If Not VALIDACGC(txtCNPJCPF.Text) Then
                         MsgBox "CGC com DV incorreto !!! "
                         txtCNPJCPF.PromptInclude = False
                         txtCNPJCPF = ""
                         txtCNPJCPF.SetFocus
                         Exit Sub
                      End If
                    Case Is > 14
                       MsgBox "CGC/CPF com DV incorreto !!! "
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    Case Is < 11
                       MsgBox "CGC/CPF com DV incorreto !!! "
                       txtCNPJCPF = ""
                       txtCNPJCPF.SetFocus
                       Exit Sub
                 End Select
                 Else
                    MsgBox "CGC/CPF com DV incorreto !!! "
                    txtCNPJCPF = ""
                    txtCNPJCPF.SetFocus
                    Exit Sub
              End If
              txtCNPJCPF.PromptInclude = False
              CRITERIO_A = txtCNPJCPF.Text
        End If
        txtCNPJCPF.PromptInclude = False
        If Trim(txtCNPJCPF.Text) <> "" Then
           CRITERIO_A = txtCNPJCPF.Text
           If Not IsNull(txtCNPJCPF.Text) Then
              If Len(txtCNPJCPF.Text) <= 11 Then
                 txtCNPJCPF.Mask = "###.###.###-##"
                 Else: txtCNPJCPF.Mask = "##.###.###/####-##"
              End If
           End If
           txtCNPJCPF.Text = CRITERIO_A
           Else: txtCNPJCPF.Mask = "##############"
        End If
        txtCNPJCPF.PromptInclude = False

         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select * from vwFornecedor "
         SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
         TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCliente.EOF Then
            FORNEC_ID_N = 0 & TabCliente!FORNECEDOR_ID
            txtRazao.Text = TabCliente!NOME
         End If
         If TabCliente.State = 1 Then _
            TabCliente.Close
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub cmdGrid_Click()
'On Error GoTo ERRO_TRATA

   Dim i As Integer

   SQL_CONSULTA = ""
   SqL2 = ""
   INDR_PRI = True

   If lstCampos.ListItems.Count > 0 Then
      For i = lstCampos.ListItems.Count To 1 Step -1
         If lstCampos.ListItems(i).Checked = True Then
            If Trim(lstCampos.ListItems(i).Text) <> "" Then

               If INDR_PRI = True Then
                  SqL2 = Trim(lstCampos.ListItems(i).Text)
                  INDR_PRI = False
                  Else: SqL2 = SqL2 & "," & Trim(lstCampos.ListItems(i).Text)
               End If

            End If
         End If
      Next i
   End If

   SQL_CONSULTA = "select " & SqL2 & " from " & Trim(cmbTabela.Text)
   SQL_CONSULTA = SQL_CONSULTA & " where empresa_id = " & EMPRESA_ID_N

   If Trim(cmbFamilia.Text) <> "" Then _
      SQL_CONSULTA = SQL_CONSULTA & " and familiaproduto_id = " & numeros(cmbFamiliaAUX.Text)

   If FORNEC_ID_N > 0 Then _
      SQL_CONSULTA = SQL_CONSULTA & " and fornecedor_id = " & FORNEC_ID_N

   If chkZerados.Value = 0 Then _
      SQL_CONSULTA = SQL_CONSULTA & " and qtde > 0 "

   If chkInativos.Value = 0 Then _
      SQL_CONSULTA = SQL_CONSULTA & " and situacao = 'A'"

   SETA_GRID SQL_CONSULTA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL_CONSULTA, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      cmdExportar.Enabled = True
      Else: cmdExportar.Enabled = False
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_LISTA_CAMPOS"
End Sub

Private Sub cmdExportar_Click()
'On Error GoTo ERRO_TRATA

   frmINICIO.Dialogo.DialogTitle = "Selecionar Caminho Arquivo!"
   frmINICIO.Dialogo.Filter = "*.txt;*.xls"
   frmINICIO.Dialogo.ShowOpen
   If Trim(frmINICIO.Dialogo.FileName) <> "" Then
      CRITERIO_A = frmINICIO.Dialogo.FileName
      Else: Exit Sub
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   TabTemp.Open SQL_CONSULTA, CONECTA_RETAGUARDA, , , adCmdText

   ExportaExcel TabTemp, frmINICIO.Dialogo.FileName

   If TabTemp.State = 1 Then _
      TabTemp.Close

   CRITERIO_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdExportar_Click"
End Sub

Sub IMPRIMIR_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{PRODUTO.TIPO_PROD} = 1"
   FORMULA_REL = FORMULA_REL & " and {PRODUTO.situacao} = 'A'"

   If cmbFamilia.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.familiaproduto_id} = " & numeros(cmbFamiliaAUX.Text)

   If FORNEC_ID_N > 0 Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.fornecedor_id} = " & FORNEC_ID_N

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
   
   Nome_Relatorio = "familia_produto.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIMIR_REL"
End Sub

Sub PreencheComboGrupo(NomeCombo As ComboBox)
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from FAMILIAPRODUTO "
   SQL = SQL & " order by descricao"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      'Mundando o ponteiro do mouse, para mostrar para o usuario que esta processando...
      Screen.MousePointer = vbHourglass

      TabTemp.MoveFirst
      Do Until TabTemp.EOF
         'Importantissimo
         DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"

         cmbFamilia.AddItem Trim(TabTemp!DESCRICAO) & "-" & TabTemp!CODG_FAMILIA
         cmbFamiliaAUX.AddItem TabTemp!FAMILIAPRODUTO_ID

         TabTemp.MoveNext
      Loop
   End If
   
   'Voltando o ponteiro do mouse para o tipo default, ponteiro.
   Screen.MousePointer = vbDefault

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "preencheComboGRUPO"
End Sub

Sub CARREGA_TABELAS()
'On Error GoTo ERRO_TRATA

   Dim rsTabela As New ADODB.Recordset

   SQL = "select name from sysobjects "
   SQL = SQL & " WHERE sysobjects.type = 'U'"
   SQL = SQL & " order by name"
   rsTabela.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not rsTabela.EOF

      cmbTabela.AddItem Trim(rsTabela.Fields(0).Value)

      rsTabela.MoveNext
   Wend
   rsTabela.Close

   cmbTabela.Text = "PRODUTO"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_TABELAS"
End Sub

Sub CARREGA_LISTA_CAMPOS(strCampo As String, strTabela As String)
'On Error GoTo ERRO_TRATA

   Dim rsCampo As New ADODB.Recordset

   lstCampos.ListItems.Clear
   CONT_N = 0

   If rsCampo.State = 1 Then _
      rsCampo.Close

   SQL = "select syscolumns.Name from sysobjects "
   SQL = SQL & " INNER JOIN syscolumns "
   SQL = SQL & " ON sysobjects.id = syscolumns.id "
   SQL = SQL & " WHERE (sysobjects.xtype = 'U') "
   SQL = SQL & " and sysobjects.name = '" & Trim(strTabela) & "'"

   'SQL = SQL & " and syscolumns.name NOT is NULL"

   If Trim(strCampo) <> "" Then _
      SQL = SQL & " and syscolumns.name = '" & Trim(strCampo) & "'"

   rsCampo.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not rsCampo.EOF

      CONT_N = CONT_N + 1
      Set item = lstCampos.ListItems.Add(, "seq." & CONT_N, Trim(rsCampo.Fields(0).Value))

      rsCampo.MoveNext
   Wend
   If rsCampo.State = 1 Then _
      rsCampo.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_LISTA_CAMPOS"
End Sub

Private Sub SETA_GRID(strSQL As String)
'On Error GoTo ERRO_TRATA

   adoCampos.Enabled = True
   adoCampos.ConnectionString = AUTENTICA_GRID
   adoCampos.RecordSource = strSQL
   adoCampos.Enabled = True
   adoCampos.Refresh

   Dim i As Integer

   CONT_N = 0

   If lstCampos.ListItems.Count > 0 Then
      For i = lstCampos.ListItems.Count To 1 Step -1
         If lstCampos.ListItems(i).Checked = True Then
            If Trim(lstCampos.ListItems(i).Text) <> "" Then

               grdCampos.Columns(CONT_N).DataField = Trim(lstCampos.ListItems(i).Text)
               grdCampos.Columns(CONT_N).Caption = Trim(lstCampos.ListItems(i).Text)
               grdCampos.Columns(CONT_N).Width = 1000

               CONT_N = 1 + CONT_N
            End If
         End If
      Next i
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub LIMPA_TUDO()
   SETA_GRID "select * from produto where produto_id < 0"
   chkZerados.Value = 1
   chkInativos.Value = 0
   cmdExportar.Enabled = False
   lstCampos.ListItems.Clear
   cmbFamilia.Text = ""
   cmbFamiliaAUX.Text = ""
   FORNEC_ID_N = 0
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtRazao.Text = ""
   CARREGA_LISTA_CAMPOS "", "PRODUTO"
End Sub
