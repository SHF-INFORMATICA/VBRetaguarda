VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmChecarDoc 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   12465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "VERIFICAR NFe/NFCe"
      Height          =   735
      Left            =   6600
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin MSComctlLib.ListView lstPedido 
      Height          =   585
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   1032
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
      ForeColor       =   4194304
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido_id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CNPJCPF"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dt_Pedido"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Prefixo"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstPedidoItem 
      Height          =   1665
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   2937
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   4194304
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IBGE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CIDADE"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UF"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CEP"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView lstFaturamento 
      Height          =   825
      Left            =   360
      TabIndex        =   3
      Top             =   4440
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   1455
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   4194304
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IBGE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CIDADE"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UF"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CEP"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmChecarDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
   Dim TabCupom      As New ADODB.Recordset
   Dim TabMFA010     As New ADODB.Recordset
   Dim TabEstab      As New ADODB.Recordset
   Dim TabPedido     As New ADODB.Recordset
   Dim TabPedidoItem As New ADODB.Recordset
   Dim Seq_Cupom     As Long
   Dim MFADOC_A      As String
   Dim ID_NF_N       As Long
   Dim CFOP_N        As String
   Dim NUMR_SEQ_A    As String
   Dim TRANSP_ID_N   As Long

   Seq_Cupom = 0
   NUMR_SEQ_N = 0

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   If TabMFA010.State = 1 Then _
      TabMFA010.Close

   SQL = " select MAX(CONVERT(INT,MFADOC)) from MFA010 "
   TabMFA010.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabMFA010.EOF Then _
      If Not IsNull(TabMFA010.Fields(0).Value) Then _
         Seq_Cupom = 0 & TabMFA010.Fields(0).Value
   If TabMFA010.State = 1 Then _
      TabMFA010.Close

NUMR_SEQ_N = 4439
Seq_Cupom = 4440

   While NUMR_SEQ_N <> Seq_Cupom
      NUMR_SEQ_N = NUMR_SEQ_N + 1

Command4.Caption = NUMR_SEQ_N
DoEvents

'MFA010
      If TabMFA010.State = 1 Then _
         TabMFA010.Close

      SQL = " select MFASEQUENCIA, MFADOC, MFACHAVENFE,MFACODMORE,MFACODSTAT,MFACODPROT "
      SQL = SQL & " from MFA010 WITH (NOLOCK) "
      SQL = SQL & " where mfadoc = '" & Trim(NUMR_SEQ_N) & "'"
      TabMFA010.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If TabMFA010.EOF Then  'não achou no mfa010

'vai ler na tabela cupom do BANCO MEGASIM
         If TabCupom.State = 1 Then _
            TabCupom.Close
         SQL = " select * from CUPOM WITH (NOLOCK) "
         SQL = SQL & " where numr_cupom = " & NUMR_SEQ_N
         TabCupom.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCupom.EOF Then  'achou no cupom
            PEDIDO_ID_N = 0 & TabCupom.Fields("pedido_id").Value
            'vai ler na tabela pedido do BANCO MEGASIM
            If TabPedido.State = 1 Then _
               TabPedido.Close

            SQL = " select * from pedido WITH (NOLOCK) "
            SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPedido.EOF Then  'achou no pedido
               If TabPedidoItem.State = 1 Then _
                  TabPedidoItem.Close

               SQL = " select * from pedidoitem WITH (NOLOCK) "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabPedidoItem.EOF Then
                  NUMR_SEQ_A = NUMR_SEQ_N
                  TRATA_CLIENTE (TabPedido.Fields("cgccpf").Value)
                  
                  GRAVA_NOTA NUMR_SEQ_A, _
                             "1", _
                             "NFC", _
                             "P", _
                             "1", _
                             "UN", _
                             "", _
                             "", _
                             "1", _
                             "1", _
                             "5102", _
                             ""
'===============================
                     If TabProduto.State = 1 Then _
                        TabProduto.Close
                     SQL = "select NFITEM.NF_ID, NFITEM.SEQ_ID, NFITEM.PRODUTO_ID"
                     SQL = SQL & " from NF "
                     SQL = SQL & " INNER JOIN NFITEM "
                     SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID"

                     SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                     TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     While Not TabProduto.EOF
                        ID_NF_N = 0 & TabProduto.Fields("nf_id").Value
                        frmINTEGRA.INTEGRA_PRODUTO (TabProduto.Fields("produto_id").Value)
                        TabProduto.MoveNext
                     Wend
                     If TabProduto.State = 1 Then _
                        TabProduto.Close

                     SqL2 = ""

                     SQL = "select distinct(cfop_id) from NFITEM"
                     SQL = SQL & " where nf_id = " & ID_NF_N
                     TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TabCliente.EOF Then
                        CFOP_N = "" & TabCliente.Fields(0).Value

                        If TabCliente.State = 1 Then _
                           TabCliente.Close

                        SQL = "select descricao from CFOP "
                        SQL = SQL & " where cfop_id = " & CFOP_N
                        TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If Not TabCliente.EOF Then _
                           SqL2 = "" & Trim(TabCliente.Fields(0).Value)
                        If TabCliente.State = 1 Then _
                           TabCliente.Close
                     End If
                     If TabCliente.State = 1 Then _
                        TabCliente.Close

                  SQL3 = "" & Trim("Tributos Totais Incidentes(Lei Federal 12.741/2012): R$ " & Format(VALOR_TOTAL_IMPOSTO_N, strFormatacao2Digitos))

                  TRANSP_ID_N = 0 & TRAZ_ID_TABELA("vwTRANSPORTADORA", "transp_id", "cnpjcpf", CNPJ_EMPRESA_N)

                  Call frmINTEGRA.INTEGRA_PEDIDO(ID_NF_N, _
                                             TRANSP_ID_N, _
                                             "", _
                                             "NFC", _
                                             "", _
                                             SQL3, _
                                             "1", _
                                             "1", _
                                             "1", _
                                             "", _
                                             "1", _
                                             "", _
                                             "", _
                                             "9", _
                                             "", _
                                             SqL2, _
                                             "1", _
                                             "0", _
                                             "0", _
                                             "0")

                  Call frmINTEGRA.INTEGRA_FINANCEIRO("NFE")
                  SQL3 = ""
                  SqL2 = ""
'==================================

               End If   'If Not TabPedidoItem.EOF Then
               If TabPedidoItem.State = 1 Then _
                  TabPedidoItem.Close
               Else  'não achou na tabela pedido
'=======================

            End If   'If TabPedido.EOF Then  'achou no pedido
            Else 'NÃO ACHOU NA TABELA CUPOM PODE SER NFe
'=======================

         End If   'If Not TabCupom.EOF Then  'achou no cupom
         If TabCupom.State = 1 Then _
            TabCupom.Close
      End If
      PEDIDO_ID_N = 0
      PESSOA_ID_N = 0
   Wend
   If TabCupom.State = 1 Then _
      TabCupom.Close
   If TabMFA010.State = 1 Then _
      TabMFA010.Close
   If TabEstab.State = 1 Then _
      TabEstab.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close
End Sub

Private Sub lixo()
   Dim TabMFA010     As New ADODB.Recordset
   Dim TabEstab      As New ADODB.Recordset
   Dim TabPedido     As New ADODB.Recordset
   Dim TabPedidoItem As New ADODB.Recordset
   Dim Seq_Cupom     As Long
   Dim MFADOC_A      As String
   Dim ID_NF_N       As Long
   Dim CFOP_N        As String
   Dim NUMR_SEQ_A    As String
   Dim TRANSP_ID_N   As Long

   Seq_Cupom = 0
   NUMR_SEQ_N = 0

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   If TabMFA010.State = 1 Then _
      TabMFA010.Close

   SQL = " select MAX(CONVERT(INT,MFADOC)) from MFA010 "
   TabMFA010.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabMFA010.EOF Then _
      If Not IsNull(TabMFA010.Fields(0).Value) Then _
         Seq_Cupom = 0 & TabMFA010.Fields(0).Value
   If TabMFA010.State = 1 Then _
      TabMFA010.Close

   While NUMR_SEQ_N < Seq_Cupom
      NUMR_SEQ_N = NUMR_SEQ_N + 1

'MFA010
      If TabMFA010.State = 1 Then _
         TabMFA010.Close

      SQL = " select MFASEQUENCIA, MFADOC, MFACHAVENFE,MFACODMORE,MFACODSTAT,MFACODPROT "
      SQL = SQL & " from MFA010 WITH (NOLOCK) "
      SQL = SQL & " where mfadoc = '" & Trim(NUMR_SEQ_N) & "'"
      TabMFA010.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If TabMFA010.EOF Then  'não achou no mfa010
         'vai ler na tabela cupom do BANCO MEGASIM
         If TabCupom.State = 1 Then _
            TabCupom.Close

         SQL = " select * from CUPOM WITH (NOLOCK) "
         SQL = SQL & " where numr_cupom = " & NUMR_SEQ_N
         TabCupom.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCupom.EOF Then  'não achou no cupom
            'vai ler na tabela nf do BANCO MEGASIM
            If TabCupom.State = 1 Then _
               TabCupom.Close

            SQL = " select * from nf WITH (NOLOCK) "
            SQL = SQL & " where numr_nota = " & NUMR_SEQ_N
            SQL = SQL & " and dt_emissao > '" & DMA("01/07/2017", "I") & "'"
            SQL = SQL & " and MODELO_DOC is not null"

'Debug.Print SQL

            TabCupom.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabCupom.EOF Then  'não achou no nf
'=================
'=================
               Else  'achou na tabela cupom vai buscar na tabela pedido
            End If
            Else  'achou na tabela cupom
               PEDIDO_ID_N = 0 & TabPedido.Fields("pedido_id").Value
               'vai ler na tabela pedido do BANCO MEGASIM
               If TabPedido.State = 1 Then _
                  TabPedido.Close

               SQL = " select * from pedido WITH (NOLOCK) "
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabPedido.EOF Then  'não achou no nf
                  Else  'achou na tabela pedido
                     If TabPedidoItem.State = 1 Then _
                        TabPedidoItem.Close

                     SQL = " select * from pedidoitem WITH (NOLOCK) "
                     SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                     TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If Not TabPedidoItem.EOF Then
                        NUMR_SEQ_A = NUMR_SEQ_N
                        
                        GRAVA_NOTA NUMR_SEQ_A, _
                                   "1", _
                                   "NFC", _
                                   "P", _
                                   "1", _
                                   "UN", _
                                   "", _
                                   "", _
                                   "1", _
                                   "1", _
                                   "5102", _
                                   ""
'===============================
                           If TabProduto.State = 1 Then _
                              TabProduto.Close
                           SQL = "select NFITEM.NF_ID, NFITEM.SEQ_ID, NFITEM.PRODUTO_ID"
                           SQL = SQL & " from NF "
                           SQL = SQL & " INNER JOIN NFITEM "
                           SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID"

                           SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                           TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           While Not TabProduto.EOF
                              ID_NF_N = 0 & TabProduto.Fields("nf_id").Value
                              frmINTEGRA.INTEGRA_PRODUTO (TabProduto.Fields("produto_id").Value)
                              TabProduto.MoveNext
                           Wend
                           If TabProduto.State = 1 Then _
                              TabProduto.Close

                           SqL2 = ""

                           SQL = "select distinct(cfop_id) from NFITEM"
                           SQL = SQL & " where nf_id = " & ID_NF_N
                           TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If Not TabCliente.EOF Then
                              CFOP_N = "" & TabCliente.Fields(0).Value

                              If TabCliente.State = 1 Then _
                                 TabCliente.Close

                              SQL = "select descricao from CFOP "
                              SQL = SQL & " where cfop_id = " & CFOP_N
                              TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                              If Not TabCliente.EOF Then _
                                 SqL2 = "" & Trim(TabCliente.Fields(0).Value)
                              If TabCliente.State = 1 Then _
                                 TabCliente.Close
                           End If
                           If TabCliente.State = 1 Then _
                              TabCliente.Close

   SQL3 = "" & Trim("Tributos Totais Incidentes(Lei Federal 12.741/2012): R$ " & Format(VALOR_TOTAL_IMPOSTO_N, strFormatacao2Digitos))

   TRANSP_ID_N = 0 & TRAZ_ID_TABELA("vwTRANSPORTADORA", "transp_id", "cnpjcpf", CNPJ_EMPRESA_N)

   Call frmINTEGRA.INTEGRA_PEDIDO(ID_NF_N, _
                                 TRANSP_ID_N, _
                                 "", _
                                 "NFC", _
                                 "", _
                                 SQL3, _
                                 "1", _
                                 "1", _
                                 "1", _
                                 "", _
                                 "1", _
                                 "", _
                                 "", _
                                 "9", _
                                 "", _
                                 SqL2, _
                                 "1", _
                                 "0", _
                                 "0", _
                                 "0")

                           Call frmINTEGRA.INTEGRA_FINANCEIRO("NFE")
SQL3 = ""
SqL2 = ""
'==================================
                     End If
                     If TabPedidoItem.State = 1 Then _
                        TabPedidoItem.Close
               End If
         End If
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close
      PEDIDO_ID_N = 0
   Wend
   If TabMFA010.State = 1 Then _
      TabMFA010.Close
   If TabEstab.State = 1 Then _
      TabEstab.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close
End Sub


