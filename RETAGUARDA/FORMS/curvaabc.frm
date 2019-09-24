VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcurvaabc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Curva ABC"
   ClientHeight    =   3225
   ClientLeft      =   3300
   ClientTop       =   1785
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "curvaabc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8850
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3720
      Picture         =   "curvaabc.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2160
      Width           =   405
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   3000
      Picture         =   "curvaabc.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   405
   End
   Begin VB.ComboBox cmbFamiliaAUX 
      BackColor       =   &H80000004&
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
      Height          =   360
      Left            =   1800
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      Height          =   240
      Left            =   6960
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox cmbProdutoAUX 
      BackColor       =   &H80000004&
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
      Height          =   360
      Left            =   3480
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbFamilia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Informe aqui o Grupo de Produtos se deseja imprimir Relatorio de Contagem por Grupo!"
      Top             =   2640
      Width           =   6855
   End
   Begin VB.TextBox txtNome 
      DataField       =   "Nome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      MaxLength       =   80
      TabIndex        =   5
      Top             =   2160
      Width           =   4335
   End
   Begin VB.ComboBox cmbProduto 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   5055
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   2
      ToolTipText     =   "Informe o código do produto."
      Top             =   1680
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1035
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   4440
      TabIndex        =   1
      Top             =   1035
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   345
      Left            =   1680
      TabIndex        =   4
      ToolTipText     =   "Informe aqui o Fornecedor se Deseja Imprimir Relatorio de Contagem por Fornecedor!"
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
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
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6600
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
               Picture         =   "curvaabc.frx":7016
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "curvaabc.frx":81B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "curvaabc.frx":923F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "curvaabc.frx":A1F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "curvaabc.frx":B414
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      Height          =   2295
      Left            =   120
      Top             =   840
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Familia:"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
      Height          =   240
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   1035
   End
End
Attribute VB_Name = "frmcurvaabc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call CentralizaJanela(frmcurvaabc)
   Me.Caption = Me.Caption & " - " & Me.Name

   PreencheComboGrp cmbFamilia, cmbFamiliaAux
   CARREGA_PRODUTO

   VALOR_TOTAL_N = 0
   PERC_ACUM_N = 0
   PERC_ACUM_VLR_N = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
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
         txtDtIni.PromptInclude = False
         txtDtFim.PromptInclude = False
         txtCNPJCPF.PromptInclude = False
         txtDtIni.Text = ""
         txtDtFim.Text = ""
         txtCodigo.Text = ""
         cmbProduto.Text = ""
         cmbProdutoAux.Text = ""
         txtCNPJCPF.Text = ""
         txtNome.Text = ""
         cmbFamilia.Text = ""
         cmbFamiliaAux.Text = ""
      Case "print"
         GERA_IMRPESSAO
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub


Private Sub cmdConsProd_Click()
   SQL3 = ""
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtCodigo.Text = SQL3
      txtCodigo.SetFocus
   End If
End Sub

Private Sub cmbProduto_Click()
On Error Resume Next

   cmbProdutoAux.ListIndex = cmbProduto.ListIndex
End Sub

Private Sub txtCNPJCPF_GotFocus()
   txtCNPJCPF.Mask = "##############"
   txtCNPJCPF.PromptInclude = True
End Sub

Private Sub txtCNPJCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_FORNEC
      txtCNPJCPF.SetFocus
   End If
End Sub

Private Sub txtCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         frmDISPLAYFORNECEDOR.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
            txtCNPJCPF.PromptInclude = True

            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close

            SQL = "select nome,razao_social from FORNECEDOR "
            SQL = SQL & " where CGCCPF = '" & CNPJCPF_A & "'"
            TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabFornecedor.EOF Then
               If Trim(TabFornecedor!NOME) = "" Then
                  txtNome.Text = Trim(TabFornecedor!razao_social)
                  Else: txtNome.Text = Trim(TabFornecedor!NOME)
               End If
            End If
            If TabFornecedor.State = 1 Then _
               TabFornecedor.Close
         End If
         CNPJCPF_A = ""
         txtCNPJCPF.SetFocus
   End Select
End Sub

Private Sub cmbFamilia_Click()
On Error Resume Next

   cmbFamiliaAux.ListIndex = cmbFamilia.ListIndex
End Sub

Private Sub txtCodigo_GotFocus()
   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consultar Produtos", "", "", ""
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmProdutoConsulta.Show 1
         If SQL3 <> "" Then
            txtCodigo.Text = SQL3
            txtCodigo.SetFocus
         End If
   End Select
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      
      If txtCodigo.Text = "" Then
         cmbProduto.SetFocus
         Else
            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select descricao from PRODUTO "
            SQL = SQL & " where codg_produto = '" & Trim(txtCodigo.Text) & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then _
               cmbProduto.Text = Trim(TabProduto!DESCRICAO)
            If TabProduto.State = 1 Then _
               TabProduto.Close

            txtDtIni.SetFocus
      End If
   End If
End Sub

Private Sub txtDTfim_GotFocus()
   txtDtFim.PromptInclude = False
   If Trim(txtDtFim.Text) = "" Then _
      txtDtFim.Text = Date
   txtDtFim.PromptInclude = True
End Sub

Private Sub txtDTINI_GotFocus()
   txtDtIni.PromptInclude = False
   If Trim(txtDtIni.Text) = "" Then _
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
End Sub

Private Sub txtDTINI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If
End Sub

Private Sub GERA_IMRPESSAO()
   If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then
      'Criando tabela para acumalar valores e percentuais para analise CURVA ABC
      If EXISTE_OBJ_BANCO("RETAGUARDA", "ABCREL", "U") = True Then
         CONECTA_RETAGUARDA.Execute "DROP table ABCREL"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      SQL = " create table ABCREL "
      SQL = SQL & "  (empresa_id int"
      SQL = SQL & " ,codg_produto nvarchar(30)"
      SQL = SQL & " ,desc_prod nvarchar(100)"
      SQL = SQL & " ,cgccpf nvarchar(16)"
      SQL = SQL & " ,qtd_acum float"
      SQL = SQL & " ,vlr_acum float"
      SQL = SQL & " ,vlr_vendido float"
      SQL = SQL & " ,vlr_venda float"
      SQL = SQL & " ,perc_acum float"
      SQL = SQL & " ,perc_acum_vlr float"
      SQL = SQL & " ,perc_total_qtd float"
      SQL = SQL & " ,perc_total_vlr float"
      SQL = SQL & " ,QTDE_PEDIDO float"
      SQL = SQL & " ,vlr_acum_venda float"
      SQL = SQL & " ,grupo int)"
      CONECTA_RETAGUARDA.Execute SQL

      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      SQL = "select * from QryFinalKardex "
      SQL = SQL & " where dt_entrada >= '" & DMA(txtDtIni.Text) & "'"
      SQL = SQL & " and dt_entrada <= '" & DMA(txtDtFim.Text) & "'"
      'SQL = SQL & " and tipo = '" & "SAIDA" & "'"
      TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabCABECA.EOF   'Dados Pedidos Validos
         If TabAUX.State = 1 Then _
            TabAUX.Close

         SQL = "select * from ABCREL "
         SQL = SQL & " where codg_proditp = '" & Trim(TabCABECA!Codg_Produto) & "'"
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabAUX.EOF Then
            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select * from PRODUTO "
            SQL = SQL & " where codg_produto = '" & Trim(TabCABECA!Codg_Produto) & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and situacao <> 'C' "
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabProduto.EOF Then
               'TabAUX!vlr_venda = TabProduto!PRECO_VENDA
               'TabAUX!vlr_acum_venda = (TabCABECA!qtde_entrada * TabProduto!PRECO_VENDA)
               'TabAUX!CGCCPF = TabProduto!CGCCPF
               'TabAUX!Grupo = TabProduto!Grupo

            SqL2 = "INSERT INTO ABCREL "
               SqL2 = SqL2 & " (Empresa_id, Codg_Produto, desc_prod, vlr_acum, "
               SqL2 = SqL2 & " qtd_acum, vlr_vendido, vlr_venda, vlr_acum_venda, CGCCPF, Grupo,"
               SqL2 = SqL2 & " perc_acum, perc_acum_vlr, perc_total_qtd, perc_total_vlr, QTDE_PEDIDO)"
            SqL2 = SqL2 & " VALUES ("
               SqL2 = SqL2 & EMPRESA_ID_N                                                          'Empresa_id
               SqL2 = SqL2 & ",'" & Trim(TabCABECA!Codg_Produto) & "'"                                'Codg_Produto
               SqL2 = SqL2 & ",'" & Trim(TabProduto.Fields("descricao").Value) & "'"               'desc_prod
               SqL2 = SqL2 & "," & tpMOEDA(TabCABECA!qtde_entrada * TabCABECA!PRECO_CUSTO)          'vlr_acum
               SqL2 = SqL2 & "," & tpMOEDA(TabCABECA!qtde_entrada)                                  'qtd_acum
               SqL2 = SqL2 & "," & tpMOEDA(TabCABECA!PRECO_CUSTO)                                  'vlr_vendido
               SqL2 = SqL2 & "," & tpMOEDA(TabProduto!PRECO_VENDA)                                 'vlr_venda
               SqL2 = SqL2 & "," & tpMOEDA((TabCABECA!qtde_entrada * TabProduto!PRECO_VENDA))       'vlr_acum_venda
               SqL2 = SqL2 & ",'" & TabCABECA.Fields("fornecedor_id").Value & "'"                  'CGCCPF
               SqL2 = SqL2 & "," & TabProduto.Fields("familiaproduto_id").Value                    'Grupo
               SqL2 = SqL2 & "," & tpMOEDA(0)                                                                 'perc_acum
               SqL2 = SqL2 & "," & tpMOEDA(0)                                                                  'perc_acum_vlr
               SqL2 = SqL2 & "," & tpMOEDA(0)                                                                  'perc_total_qtd
               SqL2 = SqL2 & "," & tpMOEDA(0)                                                                  'perc_total_vlr
               SqL2 = SqL2 & ",1"                                                                  'QTDE_PEDIDO "
            SqL2 = SqL2 & ")"

            CONECTA_RETAGUARDA.Execute SqL2
            End If
            If TabProduto.State = 1 Then _
               TabProduto.Close

            Else
               SqL2 = " Update ABCREL Set "
                  SqL2 = SqL2 & " Vlr_Acum = " & tpMOEDA((TabCABECA!qtde_entrada * TabCABECA!PRECO_CUSTO) + TabAUX!vlr_acum)
                  SqL2 = SqL2 & ", qtd_acum = " & tpMOEDA(TabCABECA!qtde_entrada + TabAUX!qtd_acum)
                  SqL2 = SqL2 & ", vlr_acum_venda = " & tpMOEDA((TabCABECA!qtde_entrada * TabAUX!vlr_venda) + TabAUX!vlr_acum_venda)
                  SqL2 = SqL2 & ", QTDE_PEDIDO = " & TabAUX!QTDE_PEDIDO + 1
               SqL2 = SqL2 & " where codg_produto = '" & TabAUX!Codg_Produto & "'"
               SqL2 = SqL2 & " and empresa_id = " & EMPRESA_ID_N
               CONECTA_RETAGUARDA.Execute SqL2
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         DoEvents

         TabCABECA.MoveNext
      Wend
      If TabCABECA.State = 1 Then _
         TabCABECA.Close

      If TabAUX.State = 1 Then _
         TabAUX.Close

      'Acumulando Percentuais
      SQL = "select * from ABCREL "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabAUX.EOF
         VALOR_TOTAL_N = 0

         If TabTemp.State = 1 Then _
            TabTemp.Close

         'Percentual quantidade
         SQL = "select sum(qtd_acum) from ABCREL "
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            VALOR_TOTAL_N = TabTemp.Fields(0).Value
            'TabAUX!perc_acum = Format(TabAUX!qtd_acum / VALOR_TOTAL_N * 100, strFormatacao3Digitos)
            PERC_ACUM_N = (TabAUX!perc_acum + PERC_ACUM_N)
            'TabAUX!perc_total_qtd = PERC_ACUM_N
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         VALOR_TOTAL_N = 0
         'Percentual quantidade
         SQL = "select sum(vlr_acum) from ABCREL "
         'SQL = SQL & " where codg_prod = '" & TABAUX!Codg_Prod & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            VALOR_TOTAL_N = TabTemp.Fields(0).Value
            'TabAUX!perc_acum_vlr = Format(TabAUX!vlr_acum / VALOR_TOTAL_N * 100, strFormatacao2Digitos)
            PERC_ACUM_VLR_N = (TabAUX!perc_acum_vlr + PERC_ACUM_VLR_N)
            'TabAUX!perc_total_vlr = PERC_ACUM_VLR_N
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "Update ABCREL Set "
         SQL = SQL & " perc_acum = " & tpMOEDA(TabAUX!qtd_acum / VALOR_TOTAL_N * 100)
         SQL = SQL & ", perc_total_qtd  = " & tpMOEDA(PERC_ACUM_N)
         SQL = SQL & ", perc_acum_vlr = " & tpMOEDA(TabAUX!vlr_acum / VALOR_TOTAL_N * 100)
         SQL = SQL & ", perc_total_vlr = " & tpMOEDA(PERC_ACUM_VLR_N)

         SQL = SQL & " where codg_produto = '" & TabAUX!Codg_Produto & "'"
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         TabAUX.MoveNext
      Wend
      If TabAUX.State = 1 Then _
         TabAUX.Close

      SQL = "select * from ABCREL "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N

      If txtCodigo.Text <> "" Then _
         SQL = SQL & " and abcrel.CODG_PRODUTO = '" & Trim(txtCodigo.Text) & "'"

      If cmbProdutoAux.Text <> "" Then _
         SQL = SQL & " and abcrel.CODG_PRODUTO ='" & cmbProdutoAux.Text & "'"

      If txtCNPJCPF.Text <> "" Then _
         SQL = SQL & " and cgccpf = '" & Trim(txtCNPJCPF.Text) & "'"

      If cmbFamilia.Text <> "" Then
         If Left(cmbFamilia.Text, 2) < 10 Then
            SQL = SQL & " and grupo = " & Left(cmbFamilia.Text, 1)
            Else: SQL = SQL & " and grupo = " & Left(cmbFamilia.Text, 2)
         End If
      End If

      TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAUX.EOF Then
         FORMULA_REL = "{ABCREL.empresa_id} = 1"

         If txtCodigo.Text <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {abcrel.CODG_PRODUTO} = '" & Trim(txtCodigo.Text) & "'"

         If txtCNPJCPF.Text <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {abcrel.cgccpf} = '" & Trim(txtCNPJCPF.Text) & "'"

         If cmbFamilia.Text <> "" Then
            If Left(cmbFamilia.Text, 2) < 10 Then
               FORMULA_REL = FORMULA_REL & " and {abcrel.grupo} = " & Left(cmbFamilia.Text, 1)
               Else: FORMULA_REL = FORMULA_REL & " and {abcrel.grupo} = " & Left(cmbFamilia.Text, 2)
            End If
         End If
      End If
      If TabAUX.State = 1 Then _
         TabAUX.Close

      If chkImp.Value = 1 Then _
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

      Nome_Relatorio = "rel_venda_abc.rpt"
      frmRELATORIO10.Show 1
   End If
End Sub

Private Function RetornaDescricaoVEND(vendId As String) As String
'On Error GoTo ERRO_TRATA

   Dim rstVEND As New ADODB.Recordset

   RetornaDescricaoVEND = ""

   SQL = "select * from VENDEDOR "
   SQL = SQL & " where vendedor_id = '" & vendId & "'"
   SQL = SQL & " order by vendedor_id "
   rstVEND.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not rstVEND.EOF Then
      rstVEND.MoveFirst
      RetornaDescricaoVEND = rstVEND!NOME_VEND
   End If
   rstVEND.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RetornaDescricaoVEND"
End Function

Private Sub PreencheComboGrp(NomeCombo As ComboBox, NomeComboAUX As ComboBox)
'On Error GoTo ERRO_TRATA

   NomeCombo.Clear
   NomeComboAUX.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from FAMILIAPRODUTO "
   SQL = SQL & " order by DESCRICAO"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      Screen.MousePointer = vbHourglass
       
      TabConsulta.MoveFirst
      Do Until TabConsulta.EOF
         NomeCombo.AddItem Trim(TabConsulta!DESCRICAO) & "-" & TabConsulta.Fields("familiaproduto_id").Value
         NomeComboAUX.AddItem TabConsulta.Fields("familiaproduto_id").Value
         TabConsulta.MoveNext
      Loop
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   Screen.MousePointer = vbDefault

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PreencheCombogrp"
End Sub

Private Sub PROCURA_FORNEC()
   txtCNPJCPF.PromptInclude = False

   SQL = "select * from FORNECEDOR "
   SQL = SQL & " where CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
   TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCliente.EOF Then
      txtNome.Text = TabCliente!NOME
   End If
   TabCliente.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_FORNEC"
End Sub

Sub CARREGA_PRODUTO()
   cmbProduto.Clear
   cmbProdutoAux.Clear

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select codg_produto,descricao from PRODUTO "
   SQL = SQL & " where situacao <> 'C' "
   SQL = SQL & " order by descricao "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      cmbProduto.AddItem Trim(TabProduto!DESCRICAO) & " - " & Trim(TabProduto!Codg_Produto)
      cmbProdutoAux.AddItem Trim(TabProduto!Codg_Produto)
      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close
End Sub
