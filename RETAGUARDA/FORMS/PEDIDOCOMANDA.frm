VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPedidoComanda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido Comanda"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PEDIDOCOMANDA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtQtde 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   14
      ToolTipText     =   "Informe a quantidade de venda deste produto."
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7680
      MaxLength       =   12
      TabIndex        =   13
      ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdConsProd 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   2520
      Picture         =   "PEDIDOCOMANDA.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   405
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3015
      MaxLength       =   29
      TabIndex        =   11
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox txtProduto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   9
      ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtItens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   9120
      TabIndex        =   2
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtTotalPedido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   9120
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtValorDig 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1111
      Left            =   9960
      Top             =   120
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   -240
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10965
      DesignHeight    =   6030
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   45
      TabIndex        =   3
      Top             =   1320
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7646
      _Version        =   393216
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar barPedido 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5655
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Picture         =   "PEDIDOCOMANDA.frx":6614
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDia 
      AutoSize        =   -1  'True
      Caption         =   "Dia"
      Height          =   240
      Left            =   9120
      TabIndex        =   17
      Top             =   1320
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   10920
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblQtde 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Qtde ="
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9240
      TabIndex        =   16
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblPreco 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Preço ="
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6960
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   10920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblProduto 
      Alignment       =   1  'Right Justify
      Caption         =   "Produto:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblItens 
      AutoSize        =   -1  'True
      Caption         =   "Itens:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9120
      TabIndex        =   7
      Top             =   4815
      Width           =   780
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9090
      TabIndex        =   6
      Top             =   2760
      Width           =   810
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11010
   End
End
Attribute VB_Name = "frmPedidoComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
   Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
   Private Const MF_BYPOSITION = &H400&

   Private LastRow            As Long ' Ultima linha em que se editou
   Private LastCol            As Long ' ultima coluna em que se editou
   Private ControlVisible     As Boolean

   Dim COMANDA_ID_N           As Long
   Dim SEQ_ID_N               As Long

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Timer1.Enabled = True

   OcultarControles
   lblMSG.Caption = NOME_EMPRESA_A
   lblMSG.Refresh

   REMOVE_MENU

   CARREGA_VENDEDOR_BALCAO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         LIMPA_TUDO
         Timer1.Enabled = False
         Unload frmPedidoComanda
      Case vbKeyF1
         If CARTAOBARRA_ID_N <= 0 Then
            PROCESSA_COMANDA "GRAVAR", "ABERTA"
         End If
         txtProduto.SetFocus
      Case vbKeyF3
         If TRAZ_TIPO_USUARIO <> 1 Then
            PROCESSA_COMANDA "LIMPAR", "EXCLUIR"
            txtProduto.SetFocus
         End If
      Case vbKeyF12
         LIMPA_TUDO
         txtProduto.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If TRAZ_TIPO_USUARIO = 1 Then _
      End

End Sub

Private Sub Timer1_Timer()
   lblDia.Caption = "" & Now
End Sub

Private Sub MSFlexGrid1_GotFocus()
   MOSTRA_RODAPE_COMANDA "ESC-Sair", "Delete-Excluir Item", "F12-LimpaTela", "", ""
End Sub

Private Sub cmdConsProd_Click()
'On Error GoTo ERRO_TRATA

   CONSULTA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsProd_Click"
End Sub

Private Sub txtDescricao_GotFocus()
   txtProduto.SetFocus
End Sub

Private Sub txtITENS_GotFocus()
'On Error GoTo ERRO_TRATA

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtITENS_GotFocus"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtQtde_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         VALOR_RECEBIDO_N = 0
         VALOR_RECEBIDO_N = 0 & InputBox(VALOR_RECEBIDO_N, "Informe Valor da Venda.")

         If Not IsNull(VALOR_RECEBIDO_N) Then
            If IsNumeric(VALOR_RECEBIDO_N) Then
               If VALOR_RECEBIDO_N > 0 Then

                  If Not IsNull(txtValor.Text) Then
                     If IsNumeric(txtValor.Text) Then
                        VALOR_ITEM_N = txtValor.Text
                        If VALOR_ITEM_N > 0 Then
                           txtQTDE.Text = VALOR_RECEBIDO_N / VALOR_ITEM_N
                           txtQTDE.Refresh
                        End If
                     End If
                  End If

               End If
            End If
         End If

   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_KeyDown"
End Sub

Private Sub txtTotalPedido_GotFocus()
'On Error GoTo ERRO_TRATA

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTotalPedido_GotFocus"
End Sub

Private Sub txtProduto_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDescricao.Enabled = False

   If UCase(Left(lblMSG.Caption, 7)) = UCase("COMANDA") Then
      MOSTRA_RODAPE_COMANDA "ESC-Sair", "F12-LimpaTela", "", "", ""
      Else: MOSTRA_RODAPE_COMANDA "F1-Incluir Comanda", "F3-Limpar Comanda", "F7-Consulta Produtos", "F12-LimpaTela", ""
   End If

   txtProduto.SelStart = 0
   txtProduto.SelLength = Len(txtProduto)
   txtProduto.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_GotFocus"
End Sub

Private Sub TXTPRODUTO_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CONSULTA_PRODUTO
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtProduto_KeyDown"
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txtProduto.ForeColor = vbBlue
   txtDescricao.ForeColor = vbBlue

   If KeyAscii = 13 Then

      KeyAscii = 0

      If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = True Then

         txtQTDE.Text = Format(QTDE_N, strFormatacao3Digitos)
         txtValor.Text = "" & Format(PR_VAREJO_N, strFormatacao2Digitos)
         CODG_PRODUTO_A = "" & Trim(TabProduto.Fields("codg_produto").Value)
         DESC_PRODUTO_A = "" & Trim(TabProduto.Fields("descricao").Value)
         txtDescricao.Text = "" & DESC_PRODUTO_A

         If INDR_LEU_POR_CODG_BARRAS = True Then
            txtQTDE.Text = 1

            Call PROCESSA_ITEM

            CODIGO_BARRAS_A = ""
            txtProduto.Enabled = True
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            txtProduto.SetFocus
            CODIGO_BARRAS_A = ""
            Exit Sub
         End If

         txtQTDE.SetFocus
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub txtQTDE_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtProduto.Text) = Empty Then
      txtProduto.SetFocus
      Exit Sub
   End If
   QTDE_N = 0 & txtQTDE.Text
   If QTDE_N <= 0 Then _
      txtQTDE.Text = 1

   txtQTDE.SelStart = 0
   txtQTDE.SelLength = Len(txtQTDE.Text)
   txtQTDE.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_GotFocus"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Len(Trim(txtQTDE.Text)) > 10 Then
         txtProduto.SetFocus
         Exit Sub
      End If
      QTDE_N = 0 & txtQTDE.Text
      If QTDE_N < 0 Then _
         txtQTDE.Text = 1

      Call PROCESSA_ITEM

      txtProduto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_KeyPress"
End Sub

Private Sub txtQtde_LostFocus()
'On Error GoTo ERRO_TRATA

   If Len(Trim(txtQTDE.Text)) >= 10 Then
      txtProduto.SetFocus
      Exit Sub
   End If

   If Trim(txtQTDE.Text) = "" Then
      txtQTDE.Text = 1
      Else
         If IsNumeric(txtQTDE.Text) Then
            QTDE_N = txtQTDE.Text
            If QTDE_N <= 0 Then _
               txtQTDE.Text = 1
         End If
   End If
   txtQTDE.Text = Format(txtQTDE.Text, strFormatacao3Digitos)
   txtQTDE.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQTDE_LostFocus"
End Sub

Private Sub txtvalor_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValor.SelStart = 0
   txtValor.SelLength = Len(txtValor.Text)
   txtValor.BackColor = &HC0FFFF
   txtQTDE.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_GotFocus"
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      VALOR_ITEM_N = 0 & txtValor.Text
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_KeyPress"
End Sub

Private Sub txtValor_LostFocus()
'On Error GoTo ERRO_TRATA

   Dim Valr_Atacado  As Double
   Dim Valr_Digitado As Double
   Dim Valr_Venda    As Double

   VALOR_ITEM_N = 0 & txtValor.Text
   If Trim(txtValor.Text) = "" Then
      txtValor.Text = Format(0, strFormatacao2Digitos)
      Else: txtValor.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If
   If VALOR_ITEM_N <= 0 Then
      txtProduto.SetFocus
      Exit Sub
      Else
         VALOR_ITEM_N = txtValor.Text
         txtValor.Text = Format(VALOR_ITEM_N, strFormatacao2Digitos)
         If VALOR_ITEM_N <= 0 Then
            MsgBox "Valor Unitário Inválido !!!"

            txtProduto.SetFocus
            Exit Sub
         End If
   End If

   txtValor.BackColor = &HFFFFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_LostFocus"
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

   OcultarControles

   ExibirCelula

   txtProduto.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 0)
   txtSeq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 11)

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2      'Editar ao pressionar F2
         ExibirCelula
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) Then
            Msg = "Deseja cancelar esse item ?  " & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
            Style = vbYesNo + 32
            Title = "Atenção."
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then
               GRAVA_COMANDA_ITEM "EXCLUIR", ATENDENTE_ID_N, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0), COMANDA_ID_N
               SETA_GRID
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_KeyDown"
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
         AtribuiValorCelula
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
      If LastCol > 3 Then
         If Not IsNumeric(txtValorDig.Text) Then
           MsgBox "Atenção Informe valores numericos !", vbInformation, "Valor Incorreto"
           Exit Sub
         End If
      End If

      AtribuiValorCelula
      'ProximaCelula
      OcultarControles

'==========ATUALIZAR GRID colunas
'3 = qtde
'4 = valor venda
'5 = desconto

      QTDE_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 3)
      VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
      SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 11)
      CODG_PRODUTO_A = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 0))

      If QTDE_N > 0 And VALOR_ITEM_N > 0 And VALOR_DESCONTO_N >= 0 And SEQ_ID_N > 0 Then
         MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * QTDE_N)), strFormatacao2Digitos)  'total item

         SQL = "update COMANDAITEM set "
         SQL = SQL & " qtde = " & tpMOEDA(QTDE_N)
         SQL = SQL & ",Valor_Item = " & tpMOEDA(VALOR_ITEM_N)

         SQL = SQL & " where comanda_id = " & COMANDA_ID_N
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         CONECTA_RETAGUARDA.Execute SQL

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
      LIMPA_BODY
      txtProduto.SetFocus
      Else
         ' ESC, cancela a edição
         If KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            txtValorDig.Visible = False
            'ControlVisible = False
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
'============================subrotinas
Private Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   lblMSG.Caption = NOME_EMPRESA_A

   If TabUSU.State = 1 Then _
      TabUSU.Close

   VALOR_ITEM_N = 0
   CARTAOBARRA_ID_N = 0
   SEQ_ID_N = 0
   PRODUTO_ID_N = 0
   PRODUTO_ID_N = 0

   MSFlexGrid1.Clear

   txtValorDig.Visible = False
   txtItens.Text = ""
   txtTotalPedido.Text = ""

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   
   LIMPA_BODY
   
   VALOR_TOTAL_N = 0
   COMANDA_ID_N = 0
   QTDE_PEDIDO = 0
   QTDE_ESTOQUE_N = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_TOTAL_DESCONTO_N = 0
   VALOR_TOTAL_N = 0
   USU_LIBERA_VENDA_N = 0
   INDR_RECEITA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0

   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtSeq.Text = ""
   txtQTDE.Text = ""

   QTDE_PEDIDO = 0
   QTDE_ESTOQUE_N = 0
   VALOR_ITEM_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_DIFERENCA_N = 0
   PRODUTO_ID_N = 0

   txtValor.Text = Format(0, strFormatacao2Digitos)
   txtQTDE.Text = Format(0, strFormatacao3Digitos)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then _
      txtProduto.Text = SQL3

   txtProduto.SetFocus
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
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
      Case 3 To 5
         texto = txtValorDig.Text

         If LastCol = 3 Then
            MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao3Digitos)
            Else: MSFlexGrid1.TextMatrix(LastRow, LastCol) = Format(texto, strFormatacao2Digitos)
         End If

         VALOR_VAREJO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
         VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, LastCol)

         If VALOR_ITEM_N < VALOR_VAREJO_N Then
            MSFlexGrid1.CellForeColor = vbRed
            MSFlexGrid1.CellFontBold = True
            MSFlexGrid1.CellBackColor = &H8000000F
            Else
               If VALOR_ITEM_N = VALOR_VAREJO_N Then
                  MSFlexGrid1.CellForeColor = vbBlack
                  MSFlexGrid1.CellFontBold = True
                  MSFlexGrid1.CellBackColor = vbCyan
                  Else
                     MSFlexGrid1.CellForeColor = vbBlue
                     MSFlexGrid1.CellFontBold = True
                     MSFlexGrid1.CellBackColor = vbWhite
               End If
         End If
      Case Else
         'texto = txtValorDig.Text
         'MSFlexGrid1.TextMatrix(LastRow, LastCol) = texto
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

Sub MOSTRA_RODAPE_COMANDA(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
'On Error GoTo ERRO_TRATA

   If Trim(Msg1) <> "" Then
      barPedido.Panels.Clear
      barPedido.Panels.Add (1)
      barPedido.Panels(1).Text = Trim(Msg1)
      barPedido.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         barPedido.Panels.Add (2)
         barPedido.Panels(2).Text = Trim(Msg2)
         barPedido.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            barPedido.Panels.Add (3)
            barPedido.Panels(3).Text = Trim(Msg3)
            barPedido.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               barPedido.Panels.Add (4)
               barPedido.Panels(4).Text = Trim(Msg4)
               barPedido.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  barPedido.Panels.Add (5)
                  barPedido.Panels(5).Text = Trim(Msg5)
                  barPedido.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_RODAPE_COMANDA"
End Sub

Private Sub REMOVE_MENU()
   Dim hMenu As Long
   hMenu = GetSystemMenu(hwnd, False)
   DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Sub CARREGA_VENDEDOR_BALCAO()
'On Error GoTo ERRO_TRATA

   VENDEDOR_ID_N = 0

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   SQL = "SELECT VENDEDOR_ID FROM PESSOA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN VENDEDOR WITH (NOLOCK)"
   SQL = SQL & " ON PESSOA.PESSOA_ID = VENDEDOR.PESSOA_ID"
   SQL = SQL & " where upper(descricao) = 'BALCAO'"
   TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then _
      If Not IsNull(TabDESCR.Fields(0).Value) Then _
         VENDEDOR_ID_N = 0 & Trim(TabDESCR.Fields(0).Value)
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_VENDEDOR_BALCAO"
End Sub

Sub PROCESSA_COMANDA(Tipo_Chamada As String, SITUACAO_A As String)
'On Error GoTo ERRO_TRATA

'ESSA ROTINA TEM O PROPÓSITO SOMENTE DE ATUALIZAR A TABELA COMANDA CABEÇA

   frmCOMANDA.Show 1

   If CARTAOBARRA_ID_N > 0 Then
      Dim TabCOMANDA       As New ADODB.Recordset
      Dim TabComandaItem   As New ADODB.Recordset

      COMANDA_ID_N = 0

      If TabCOMANDA.State = 1 Then _
         TabCOMANDA.Close
      'procurando registro comanda POR NUMERO DA COMANDA INFORMADA
      SQL = "select comanda_id,situacao from COMANDA WITH (NOLOCK)"
      SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N
      SQL = SQL & " and upper(situacao) = 'ABERTA'"
      TabCOMANDA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCOMANDA.EOF Then
         COMANDA_ID_N = 0 & TabCOMANDA.Fields("comanda_id").Value
         SITUACAO_A = "" & TabCOMANDA.Fields("situacao").Value

         If Trim(Tipo_Chamada) = "GRAVAR" Then
            GRAVA_CABECA_COMANDA "ABERTA", COMANDA_ID_N
            Else
               If Trim(Tipo_Chamada) = "EXCLUIR" Then
                  GRAVA_COMANDA_ITEM "EXCLUIR", ATENDENTE_ID_N, 0, COMANDA_ID_N
                  GRAVA_CABECA_COMANDA "EXCLUIR", COMANDA_ID_N
               End If
         End If
         Else: GRAVA_CABECA_COMANDA "ABERTA", COMANDA_ID_N
      End If
      If TabCOMANDA.State = 1 Then _
         TabCOMANDA.Close

      SETA_GRID
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_COMANDA"
End Sub

Public Sub GRAVA_CABECA_COMANDA(SITUACAO_A As String, NUMR_COMANDA_ID_N As Long)
'On Error GoTo ERRO_TRATA

   Dim TabCOMANDA    As New ADODB.Recordset

   If Trim(SITUACAO_A) = "EXCLUIR" Then
      Acao_N = 3
      'NUMR_COMANDA_ID_N = 0
      Else
         'PROCURANDO REGISTRO NA TABELA COMANDA
         If TabCOMANDA.State = 1 Then _
            TabCOMANDA.Close

         SQL = "select comanda_id from COMANDA WITH (NOLOCK)"
         SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N
         'SQL = SQL & " and upper(situacao) = '" & Trim(SITUACAO_A) & "'"
         TabCOMANDA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCOMANDA.EOF Then
            NUMR_COMANDA_ID_N = 0 & TabCOMANDA.Fields("comanda_id").Value
            Acao_N = 2
            Else
               Acao_N = 1
               NUMR_COMANDA_ID_N = 0 & MAX_ID("comanda_id", "comanda", "", "", "", "")
         End If
         If TabCOMANDA.State = 1 Then _
            TabCOMANDA.Close
   End If

   SQL = "spCOMANDA " & Acao_N & "," & NUMR_COMANDA_ID_N & "," & CARTAOBARRA_ID_N & "," & USUARIO_ID_N & ",'" & Now & "','" & SITUACAO_A & "'"
   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA_COMANDA"
End Sub

Public Sub GRAVA_COMANDA_ITEM(SITUACAO_A As String, ATENDENTE_ID_N As Long, Sequencia_N As Long, NUMR_COMANDA_ID_N As Long)
'On Error GoTo ERRO_TRATA

   Dim TabComandaItem   As New ADODB.Recordset

   If Trim(SITUACAO_A) = "EXCLUIR" Then
      Acao_N = 3
      PRODUTO_ID_N = 0
      Sequencia_N = 0
      ATENDENTE_ID_N = 0
      Else
         If Trim(txtSeq.Text) = "" Then
            Sequencia_N = 0 & MAX_ID("seq_id", "comandaitem", "comanda_id", Str(NUMR_COMANDA_ID_N), "", "")
            txtSeq.Text = "" & Sequencia_N
         End If

         'PROCURANDO REGISTRO NA TABELA COMANDAitem
         If TabComandaItem.State = 1 Then _
            TabComandaItem.Close

         SQL = "select comanda_id,seq_id,produto_id from COMANDAITEM WITH (NOLOCK)"
         SQL = SQL & " where comanda_id = " & NUMR_COMANDA_ID_N
         SQL = SQL & " and seq_id = " & Sequencia_N
         TabComandaItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabComandaItem.EOF Then
            Sequencia_N = 0 & TabComandaItem.Fields("comanda_id").Value
            PRODUTO_ID_N = 0 & TabComandaItem.Fields("produto_id").Value
            Acao_N = 2
            Else
               Acao_N = 1
               Sequencia_N = 0 & MAX_ID("seq_id", "comandaitem", "comanda_id", Str(NUMR_COMANDA_ID_N), "", "")
         End If
         If TabComandaItem.State = 1 Then _
            TabComandaItem.Close
   End If

   SQL = "spCOMANDAITEM " & Acao_N & "," & NUMR_COMANDA_ID_N & "," & Sequencia_N & "," & PRODUTO_ID_N & ",'" & tpMOEDA(QTDE_N) & "','" & tpMOEDA(VALOR_ITEM_N) & "','" & SITUACAO_A & "'," & ATENDENTE_ID_N
   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_COMANDA_ITEM"
End Sub

Sub PROCESSA_ITEM()
'On Error GoTo ERRO_TRATA

   If CARTAOBARRA_ID_N <= 0 Then
      MsgBox "Informar Comanda antes de prosseguir !!!"
      Exit Sub
   End If

   If Trim(txtQTDE.Text) <> "" Then
      If IsNumeric(txtQTDE.Text) Then
         QTDE_N = txtQTDE.Text
         If QTDE_N > 99 Then
            Msg = "Atenção quantidade informada muito alta, deseja continuar ???? !!!"
            PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
            If RESPOSTA = vbNo Then
               txtProduto.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If

   If Trim(CARTAOBARRA_ID_N) <= 0 Then
      MsgBox "Registre a comanda antes dos produtos."
      Exit Sub
   End If

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Informe codigo de Produto.", vbOKOnly, "Atenção."
      txtProduto.SetFocus
      Exit Sub
   End If

   If Not IsNull(txtValor.Text) Then
      VALOR_ITEM_N = 0 & txtValor.Text
      If VALOR_ITEM_N <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção."

         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtQTDE.Text) = "" Then
      Beep
      MsgBox "Informe a quantidade.", vbOKOnly, "Atenção."
      txtQTDE.SetFocus
      Exit Sub
      Else
         'quantidade pedida
         QTDE_PEDIDO = txtQTDE.Text
         txtQTDE.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)

         GRAVA_CABECA_COMANDA "ABERTA", COMANDA_ID_N
         GRAVA_COMANDA_ITEM "ABERTA", ATENDENTE_ID_N, SEQ_ID_N, COMANDA_ID_N
   End If

   'valor venda item
   VALOR_ITEM_N = txtValor.Text
   VALOR_TOTAL_DESCONTO_N = 0

   'valor total da Pedido, o desconto é armazenado no seu devido lugar, não entra no calculo do campo total da venda
   VALOR_TOTAL_N = VALOR_TOTAL_N + (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N

   LIMPA_BODY
   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_ITEM"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If CARTAOBARRA_ID_N <= 0 Then _
      Exit Sub

   Dim TabGrid                As New ADODB.Recordset
   Dim Coluna, Linha, Largura_Campo
   Dim VALOR_ITENS_PRODUCAO   As Double
   Dim VALOR_ITENS_REVENDA    As Double

   CONT_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0
   VALOR_TOTAL_N = 0
   VALOR_ITENS_PRODUCAO = 0
   VALOR_ITENS_REVENDA = 0

   txtItens.Text = "" & CONT_N
   txtTotalPedido.Text = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   txtTotalPedido.Text = Format(VALOR_TOTAL_N, "currency")

   MSFlexGrid1.Clear
   MSFlexGrid1.Visible = False
   MSFlexGrid1.Gridlines = flexGridFlat
   MSFlexGrid1.FixedRows = 1
   MSFlexGrid1.FixedCols = 1
   MSFlexGrid1.ScrollBars = flexScrollBarBoth
   MSFlexGrid1.AllowUserResizing = flexResizeColumns

   'MSFlexGrid1.Cols = 19                  ' Número de colunas(incluindo o cabecalho)
   'MSFlexGrid1.Rows = 2                   ' Número de linhas(com cabecalho)

   If TabGrid.State = 1 Then _
      TabGrid.Close

   SQL = "SELECT COMANDAITEM.SEQ_ID as Sq, PRODUTO.CODG_PRODUTO as Código, PRODUTO.DESCRICAO as Descrição, "
   SQL = SQL & " COMANDAITEM.QTDE as QtdeItem, COMANDAITEM.VALOR_ITEM as ValorItem, "
   SQL = SQL & " (COMANDAITEM.QTDE*COMANDAITEM.VALOR_ITEM) as TotalItem "
   SQL = SQL & " from COMANDA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN COMANDAITEM  WITH (NOLOCK)"
   SQL = SQL & " ON COMANDA.COMANDA_ID = COMANDAITEM.COMANDA_ID "
   SQL = SQL & " INNER JOIN PRODUTO  WITH (NOLOCK)"
   SQL = SQL & " ON COMANDAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N
   SQL = SQL & " and upper(COMANDA.situacao) = 'ABERTA'"

   SQL = SQL & " order by seq_id desc"

   TabGrid.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGrid.EOF Then
      lblMSG.Caption = "Comanda: " & CARTAOBARRA_ID_N
      lblMSG.Refresh

      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
      'MSFlexGrid1.FixedRows = 3
      MSFlexGrid1.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = TabGrid.Fields.Count

      ReDim largura_coluna(0 To TabGrid.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabGrid.Fields.Count - 1
         MSFlexGrid1.TextMatrix(0, Coluna) = Trim(TabGrid.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabGrid.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabGrid.EOF
         INDR_PRI = False
         'If Not IsNull(TabGrid.Fields("producao").Value) Then _
            If TabGrid.Fields("producao").Value = True Then _
               INDR_PRI = True

'=======totais
         CONT_N = CONT_N + 1
         VALOR_ITEM_N = VALOR_ITEM_N + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtdeitem").Value)

         'If Not IsNull(TabGrid.Fields("desconto").Value) Then _
            VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabGrid.Fields("desconto").Value

         If INDR_PRI = True Then
            'VALOR_ITENS_PRODUCAO = 0
            VALOR_ITENS_PRODUCAO = VALOR_ITENS_PRODUCAO + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtdeitem").Value)
            Else
               'VALOR_ITENS_REVENDA = 0
               VALOR_ITENS_REVENDA = VALOR_ITENS_REVENDA + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtdeitem").Value)
         End If
'========= verificando se o produto é de produção
'=========

         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabGrid.Fields.Count - 1
            'If Coluna = 3 Or Coluna = 7 Then
            If Coluna = 3 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGrid.Fields(Coluna).Value, strFormatacao3Digitos)
               Else
                  'If Coluna = 4 Or Coluna = 5 Or Coluna = 6 Or Coluna = 7 Or Coluna = 8 Or Coluna = 9 Or Coluna = 10 Then
                  If Coluna = 4 Or Coluna = 5 Or Coluna = 6 Then
                     MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGrid.Fields(Coluna).Value, strFormatacao2Digitos)
                     Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabGrid.Fields(Coluna).Value)
                  End If
            End If
'=========se o produto for de produção pintar linha
            If INDR_PRI = True Then
               MSFlexGrid1.Row = Linha
               MSFlexGrid1.Col = Coluna
               'flex_tst.Text = "Bold Font"
               'flex_tst.CellFontBold = True
               'flex_tst.CellForeColor = vbRed
               MSFlexGrid1.CellForeColor = &H4000&   '&H40&
            End If
'=========

            ' verifica o tamanho dos campos
            If Not IsNull(TabGrid.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabGrid.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo

         Next Coluna

'=========totais

'--------------
         TabGrid.MoveNext
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

'Sequencia
      MSFlexGrid1.ColWidth(0) = 800
      MSFlexGrid1.ColAlignment(0) = 0

'Codigo Produto
      MSFlexGrid1.ColWidth(1) = 1700
      MSFlexGrid1.ColAlignment(1) = 0

'Descrição Produto
      MSFlexGrid1.ColWidth(2) = 7000
      MSFlexGrid1.ColAlignment(2) = 0

'QTDE
      MSFlexGrid1.ColWidth(3) = 2111
      MSFlexGrid1.ColAlignment(3) = 7

'Valor Item
      MSFlexGrid1.ColWidth(4) = 2111
      MSFlexGrid1.ColAlignment(4) = 7

'Total Item
      MSFlexGrid1.ColWidth(5) = 2111
      MSFlexGrid1.ColAlignment(5) = 7
   End If

   ' fecha o recordset e a conexao
   If TabGrid.State = 1 Then _
      TabGrid.Close

   txtItens.Text = "" & CONT_N

   VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N

   txtTotalPedido.Text = "" & Format(VALOR_TOTAL_N, strFormatacao2Digitos)
   DoEvents

MSFlexGrid1.Visible = True
   VALOR_ITENS_PRODUCAO = 0
   VALOR_ITENS_REVENDA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub
