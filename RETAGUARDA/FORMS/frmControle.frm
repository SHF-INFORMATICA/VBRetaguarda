VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmControle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Recebimento"
   ClientHeight    =   2160
   ClientLeft      =   4200
   ClientTop       =   2565
   ClientWidth     =   6855
   Icon            =   "frmControle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6855
   Begin VB.CommandButton cmbSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5880
      Picture         =   "frmControle.frx":47C4A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1150
      Width           =   900
   End
   Begin VB.CommandButton cmbConfirma 
      Caption         =   "Confirma"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5880
      Picture         =   "frmControle.frx":49337
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   200
      Width           =   900
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   975
      Left            =   45
      TabIndex        =   4
      Top             =   1080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      _Version        =   262144
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Selecione a Forma de Pagamento"
      Begin VB.ComboBox cmbforma 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   2
         Top             =   400
         Width           =   3555
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pagamento:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   1695
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   855
      Left            =   45
      TabIndex        =   6
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      _Version        =   262144
      Caption         =   "Periodo a Importar"
      Begin Threed.SSFrame SSFrame3 
         Height          =   855
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1508
         _Version        =   262144
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Período a Importar"
         Begin MSMask.MaskEdBox txtDtIni 
            Height          =   375
            Left            =   1560
            TabIndex        =   0
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDtFim 
            Height          =   375
            Left            =   4080
            TabIndex        =   1
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Caption         =   "Data Final:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3120
            TabIndex        =   9
            Top             =   400
            Width           =   945
         End
         Begin VB.Label Label2 
            Caption         =   "Data Inicial:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   8
            Top             =   400
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frmControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbConfirma_Click()
     Msg = "Deseja Realmente Fazer esta Copia?"
     PERGUNTA Msg, vbYesNo + 32, "Controle de Pagamento", "DEMO.HLP", 1000
     If RESPOSTA = vbYes Then
        COPIADADOS
     End If
End Sub

Private Sub cmbFORMA_Click()
     'cmbConfirma.SetFocus
End Sub

Private Sub cmbSair_Click()
     Unload Me
End Sub

Private Sub Form_Load()
     Call CentralizaJanela(frmControle)
     preenchePgto
     txtDtIni.Text = Date
     txtDtFim.Text = Date
End Sub

Private Sub preenchePgto()
'On Error GoTo ERRO_TRATA
    Dim rstCONT As New ADODB.Recordset
    SQL = "select * from TIPOVENDA  order by TIPOVENDA_ID"
    rstCONT.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
    cmbforma.Clear
    If Not rstCONT.EOF Then
       'Mundando o ponteiro do mouse, para mostrar para o usuario que esta processando...
        Screen.MousePointer = vbHourglass
        rstCONT.MoveFirst
        Do Until rstCONT.EOF
            'Importantissimo
            DoEvents 'Libera o computador equanto o sistema trabalha. Não deixa a tela "congelar"
            cmbforma.AddItem rstCONT!tipovenda_id & "-" & rstCONT!Descricao
            rstCONT.MoveNext
        Loop
    End If
    'Voltando o ponteiro do mouse para o tipo default, ponteiro.
    Screen.MousePointer = vbDefault
    rstCONT.Close
ERRO_TRATA:
    'Fazer tratamento de erro
    Exit Sub
End Sub

Private Sub COPIADADOS()
     'Copiando Vendas
     'Lendo Banco de Dados Principal
     SQL = "Select * From PEDIDO Where EMPRESA_ID = " & EMPRESA_ID_N & " and STATUS = 5 and DT_REQ >= '" & DMA(txtDtIni.Text) & "' and DT_REQ <= '" & DMA(txtDtFim.Text) & "'"  ' and TipoVenda_id <> " & Left(cmbforma.Text, 1)
     TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
     Screen.MousePointer = vbHourglass
     Do While Not TabTemp.EOF
        DoEvents
        'insert banco auxiliar PEDIDO
        SqL2 = "INSERT INTO PEDIDO (Empresa_id, numr_req, CGCCPF, vendedor_id, Dt_Req, Nome_Cliente, Status, Tipo_Registro,Codg_USU, TIPOVENDA_ID, TIPO_DOC, CODGFUNC, VALOR_DESCONTO, PERC_DESC, VALOR_RECEBIDO, VALOR_TOTAL) "
        SqL2 = SqL2 & " VALUES (" & EMPRESA_ID_N & "," & TabTemp!numr_req & ",'" & TabTemp!CGCCPF & TabTemp!VENDEDOR_ID & ",'" & DMA(TabTemp!DT_REQ) & "','" & TabTemp!NOME_CLIENTE & "'," & TabTemp!Status & ",'" & TabTemp!TIPO_REGISTRO & "'," & TabTemp!CODG_USU & "," & TabTemp!tipovenda_id & ",'" & TabTemp!TIPO_DOC & "'," & TabTemp!codgfunc & "," & Str(TabTemp!Valor_Desconto) & "," & Str(TabTemp!PERC_desc) & "," & Str(TabTemp!VALOR_RECEBIDO) & "," & Str(TabTemp!Valor_Total) & ")"
        db.Execute SqL2
        
        'Copiando Itens da Venda da selecao
        TABITEM.Open "Select * FROM PEDIDOITEM Where Numr_Req = " & TabTemp!numr_req, CONECTA_RETAGUARDA, , , adCmdText
        Do While Not TABITEM.EOF
           SqL2 = "INSERT INTO PEDIDOITEM "
           SqL2 = SqL2 & " (EMPRESA_ID, NUMR_REQ, CODG_PROD, SEQ, QTD_PEDIDA, VALOR_ITEM, PERC_DESC, CFOP, STRIBUTARIA, VLRBASEICMS, "
           SqL2 = SqL2 & " PERCICMS, VLRICMS, VLRBASEICMSSUBST, PERCICMSSUBST, VLRICMSSUBST, PERCREDUCAOICMS, PERCIVA, PERC_IPI, VLR_IPI) "
           SqL2 = SqL2 & " VALUES ("
           SqL2 = SqL2 & EMPRESA_ID_N
           SqL2 = SqL2 & "," & TabTemp!numr_req
           SqL2 = SqL2 & ",'" & TabTemp!CGCCPF & TabTemp!VENDEDOR_ID & ",'" & DMA(TabTemp!DT_REQ) & "','" & TabTemp!NOME_CLIENTE & "'," & TabTemp!Status & ",'" & TabTemp!TIPO_REGISTRO & "'," & TabTemp!CODG_USU & "," & TabTemp!tipovenda_id & ",'" & TabTemp!TIPO_DOC & "'," & TabTemp!codgfunc & "," & Str(TabTemp!Valor_Desconto) & "," & Str(TabTemp!PERC_desc) & "," & Str(TabTemp!VALOR_RECEBIDO) & "," & Str(TabTemp!Valor_Total) & ")"
           db.Execute SqL2
           TABITEM.MoveNext
        Loop
        TABITEM.Close
        'copiando Lancamentos Financeiro da Selecao
        TABITEM.Open "Select * From LANCAMENTO Where Numr_doc = " & TabTemp!numr_req, CONECTA_RETAGUARDA, , , adCmdText
        If Not TABITEM.EOF Then
           'insert banco auxiliar LANCAMENTO
           
           'Copiando Itens do Lancamento
           TabLancamento.Open "Select * From ITEMLANCAMENTO Where Numr_doc = " & TABITEM!NUMR_DOC, CONECTA_RETAGUARDA, , , adCmdText
           If Not TabLancamento.EOF Then
              'insert banco auxiliar itemlancamento
           End If
           TabLancamento.Close
        End If
        TABITEM.Close
        TabTemp.MoveNext
     Loop
     Screen.MousePointer = vbDefault
     TabTemp.Close
     
     
     'Deletar Arquivos de acordo com a selecao feita pelo usuario da modalidade de pgto, o que o usuario selecionar e o registro que vai ficar no banco
     SQL = "Select * From PEDIDO Where EMPRESA_ID = " & EMPRESA_ID_N & " and DT_REQ >= '" & DMA(txtDtIni.Text) & "' and DT_REQ <= '" & DMA(txtDtFim.Text) & "'" And TabTemp!tipovenda_id <> " & Left(cmbforma.Text, 1)"
     TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
     Screen.MousePointer = vbHourglass
     Do While Not TabTemp.EOF
        DoEvents
        CONECTA_RETAGUARDA.Execute "Delete from PEDIDO Where Numr_req = " & TabTemp!numr_req
        CONECTA_RETAGUARDA.Execute "Delete from ITENSREQ Where Numr_req = " & TabTemp!numr_req
        CONECTA_RETAGUARDA.Execute "Delete from LANCAMENTO Where Numr_doc = " & TabTemp!numr_req
        CONECTA_RETAGUARDA.Execute "Delete from ITEMLANCANCAMENTO Where Numr_doc = " & TabTemp!numr_req
        TabTemp.MoveNext
     Loop
     Screen.MousePointer = vbDefault
     TabTemp.Close
     
     'Copiando Cadastros Cliente Funcionario
     'Lendo Banco de Dados Principal
     SQL = "Select * From CLIENTE Where EMPRESA_ID = " & EMPRESA_ID_N
     TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
     Screen.MousePointer = vbHourglass
     Do While Not TabTemp.EOF
        DoEvents
        'Copiando Clientes Cadastrados no Periodo
        
        TABITEM.Open "Select * From FUNCIONARIOCONVENIO Where CNPJEMPRESA = '" & TabTemp!CPFCGC & "'", CONECTA_RETAGUARDA, , , adCmdText
        If Not TABITEM.EOF Then
           'Copiando Funcionario Cadastrados no Periodo
           
        End If
        TABITEM.Close
        TabTemp.MoveNext
     Loop
     Screen.MousePointer = vbDefault
     TabTemp.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub txtDTINI_GotFocus()
   txtDtIni.PromptInclude = True
End Sub

Private Sub txtDtIni_LostFocus()
'On Error GoTo ERRO_TRATA
   txtDtIni.PromptInclude = True
   If Not IsDate(txtDtIni.Text) Then
      txtDtIni.PromptInclude = False
         txtDtIni.Text = Date
      txtDtIni.PromptInclude = True
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtIni_LostFocus"
End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
End Sub

Private Sub txtDTfim_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If
End Sub

