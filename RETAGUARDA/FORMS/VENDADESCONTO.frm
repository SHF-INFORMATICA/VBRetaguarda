VERSION 5.00
Begin VB.Form frmVENDADESCONTO 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Valor ou Percentual de Desconto"
   ClientHeight    =   2325
   ClientLeft      =   3180
   ClientTop       =   3345
   ClientWidth     =   7980
   Icon            =   "VENDADESCONTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1373.686
   ScaleMode       =   0  'User
   ScaleWidth      =   7492.792
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   53
      TabIndex        =   9
      Top             =   120
      Width           =   6855
      Begin VB.OptionButton optPerc 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   600
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optValor 
         Caption         =   "R$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   600
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtValorPermitido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtPerc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtToTVenda 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtValrDesc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblDescontoRevenda 
         AutoSize        =   -1  'True
         Caption         =   "Desconto não Permitido para produtos de revenda = "
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
         Height          =   225
         Left            =   90
         TabIndex        =   13
         Top             =   1680
         Width           =   4470
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Venda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total com Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   1215
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observação:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   53
      TabIndex        =   8
      Top             =   3240
      Width           =   8655
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   780
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   8415
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   7005
      Picture         =   "VENDADESCONTO.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   180
      Width           =   959
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   7005
      Picture         =   "VENDADESCONTO.frx":6D8D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1170
      Width           =   959
   End
End
Attribute VB_Name = "frmVENDADESCONTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_DESCONTO_N = 0
   PERC_DESCONTO_N = 0
   PERC_DESCONTO_USUARIO_N = 0
   CRITERIO_A = ""

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select sum(PEDIDOITEM.QTD_PEDIDA*PEDIDOITEM.VALOR_ITEM) from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then _
      VALOR_TOTAL_N = 0 & TabConsulta.Fields(0).Value
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   txtToTVenda.Text = Format(VALOR_TOTAL_N, strFormatacao2Digitos)

   optPerc.Visible = False
   optPerc.Value = False
   txtPerc.Visible = False
   optValor.Visible = True

   If INDR_LiberaPercDesconto = True Then
      optPerc.Visible = True
      txtPerc.Visible = True
      optPerc.Value = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Activate()

   If VALIDA_DESCONTO = False Then _
      Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close

   VALOR_TOTAL_DESCONTO_N = 0 & txtValrDesc.Text
   PERC_DESCONTO_N = 0 & txtPerc.Text

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub optPerc_Click()
On Error Resume Next

   txtPerc.Enabled = True
   txtPerc.Visible = True
   txtPerc.SetFocus
End Sub

Private Sub optValor_Click()
   txtValor.Visible = True
   txtValor.Enabled = True
   txtPerc.Text = ""
   txtValor.SetFocus
End Sub

Private Sub txtTot_GotFocus()
'On Error GoTo ERRO_TRATA

   txtPerc.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTot_GotFocus"
End Sub

Private Sub txtToTVenda_GotFocus()
'On Error GoTo ERRO_TRATA

   txtPerc.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtToTVenda_GotFocus"
End Sub

Private Sub txtValor_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyUp: txtPerc.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_KeyUp"
End Sub

Private Sub txtValrDesc_GotFocus()
'On Error GoTo ERRO_TRATA

   VALOR_TOTAL_DESCONTO_N = 0
   PERC_DESCONTO_N = 0
   txtPerc.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValrDesc_GotFocus"
End Sub

Private Sub cmdCancel_Click()
'On Error GoTo ERRO_TRATA

   INDR_DESCONTO_AUTORIZADO = False
   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdCancel_Click"
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = Asc(".") Then _
      KeyAscii = Asc(",")

   If KeyAscii = 13 Then
      If Trim(txtValor.Text) <> "" Then
         KeyAscii = 0
         PERC_DESCONTO_N = 0
         VALOR_TOTAL_DESCONTO_N = 0
         VALOR_TOTAL_DESCONTO_N = txtValor.Text
         VALOR_TOTAL_N = txtToTVenda.Text
         VALOR_TOTAL_venda_N = txtToTVenda.Text

         If Trim(txtValorPermitido.Text) <> "" Then _
            VALOR_TOTAL_N = txtValorPermitido.Text

         txtValrDesc.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
         txtTot.Text = Format(VALOR_TOTAL_venda_N - VALOR_TOTAL_DESCONTO_N, strFormatacao4Digitos)
         If optPerc.Value = True Then _
            txtPerc.Text = Format(((VALOR_TOTAL_DESCONTO_N / VALOR_TOTAL_N) * 100), strFormatacao4Digitos)

         Call txtPerc_KeyPress(13)
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValor_KeyPress"
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = Asc(".") Then _
      KeyAscii = Asc(",")

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtPerc.Text) = "" Then
         txtValor.SetFocus
         Else
            PERC_DESCONTO_N = 0
            VALOR_TOTAL_DESCONTO_N = 0
            PERC_DESCONTO_N = txtPerc.Text
            VALOR_TOTAL_venda_N = txtToTVenda.Text

            If Trim(txtValorPermitido.Text) <> "" Then _
               VALOR_TOTAL_N = txtValorPermitido.Text

            If optPerc.Visible = True Then
               VALOR_TOTAL_DESCONTO_N = (VALOR_TOTAL_N * PERC_DESCONTO_N / 100)
               Else: VALOR_TOTAL_DESCONTO_N = txtValor.Text
            End If
            txtValrDesc.Text = Format(VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
            txtTot.Text = Format(VALOR_TOTAL_venda_N - VALOR_TOTAL_DESCONTO_N, strFormatacao2Digitos)
            VALOR_TOTAL_N = txtToTVenda.Text

            cmdOk.SetFocus
      End If
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else
               If KeyAscii < 48 Or KeyAscii > 57 Then
                  KeyAscii = 0
               End If
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtPerc_KeyPress"
End Sub

Private Sub cmdOK_Click()
'On Error GoTo ERRO_TRATA

   If PERC_DESCONTO_N > PERC_DESCONTO_USUARIO_N Then
      Msg = "Limite de desconto ultrapassado, deseja liberar com senha superior ?"
      PERGUNTA Msg, vbYesNo + 32, "Desconto", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         frmSenha.Show 1

         If TabUSU.State = 1 Then _
            TabUSU.Close

         SQL = "select * from USUARIO WITH (NOLOCK)"
         SQL = SQL & " where senha = '" & Trim(CRITERIO_A) & "'"
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            If IsNull(TabUSU.Fields("tipo").Value) Then
               MsgBox "Não permitido, tipo usuário não informado."
               INDR_DESCONTO_AUTORIZADO = False
               Exit Sub
            End If

            If TabUSU.Fields("tipo").Value >= 4 And TabUSU.Fields("tipo").Value <= 5 Then
               Else
                  MsgBox "Não permitido, faixa de desconto não cadastrada para este usuário."
                  INDR_DESCONTO_AUTORIZADO = False
                  Exit Sub
            End If

            USU_LIBERA_VENDA_N = TabUSU.Fields("usuario_id").Value
            Else
               MsgBox "Não permitido."
               INDR_DESCONTO_AUTORIZADO = False
               Exit Sub
         End If
         If TabUSU.State = 1 Then _
            TabUSU.Close
         Else
            VALOR_TOTAL_DESCONTO_N = 0
            PERC_DESCONTO_N = 0
            INDR_DESCONTO_AUTORIZADO = False
            Exit Sub
      End If
   End If

   CRITERIO_A = ""
   SQL = "UPDATE PEDIDO SET "
   SQL = SQL & " Valor_desconto = " & tpMOEDA(VALOR_TOTAL_DESCONTO_N)
   SQL = SQL & " , Perc_desc = " & tpMOEDA(PERC_DESCONTO_N)
   SQL = SQL & " , valor_total = " & tpMOEDA(txtToTVenda.Text)
   SQL = SQL & " , valor_recebido = 0"

   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

   Unload Me

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdOK_Click"
End Sub

Function VALIDA_DESCONTO() As Boolean
'On Error GoTo ERRO_TRATA

   VALIDA_DESCONTO = False
   INDR_DESCONTO_AUTORIZADO = False

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select perc_desconto from USUARIO "
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      If Not IsNull(TabUSU.Fields(0).Value) Then _
         PERC_DESCONTO_USUARIO_N = TabUSU.Fields(0).Value
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If PERC_DESCONTO_USUARIO_N <= 0 Then
      MsgBox "Permissão para desconto não concedida !!!"
      Exit Function
   End If

   optValor.Value = False
   txtValorPermitido.Visible = False
   lblDescontoRevenda.Visible = False

   VALIDA_DESCONTO = True
   INDR_DESCONTO_AUTORIZADO = True

   PESQUISA_PEDIDO

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VALIDA_DESCONTO"
End Function

Sub PESQUISA_PEDIDO()
'On Error GoTo ERRO_TRATA

   Dim VALOR_PERMITE_DESCONTO_N As Double

   VALOR_PERMITE_DESCONTO_N = 0
   If PEDIDO_ID_N > 0 Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, "
      SQL = SQL & " PEDIDOITEM.STATUS, PRODUTO.TIPO_PROD, permite_desconto "
      SQL = SQL & " from PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
      SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and permite_desconto = 1 "
      SQL = SQL & " and pedidoitem.status <> 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF

         VALOR_PERMITE_DESCONTO_N = VALOR_PERMITE_DESCONTO_N + (TabConsulta.Fields("valor_item").Value * TabConsulta.Fields("QTD_PEDIDA").Value)
         txtValorPermitido.Text = Format(VALOR_PERMITE_DESCONTO_N, strFormatacao2Digitos)
         txtValorPermitido.Refresh
         txtValorPermitido.Visible = True
         optValor.Visible = True
         optValor.Value = True
         lblDescontoRevenda.Visible = True

         DoEvents

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PESQUISA_PEDIDO"
End Sub
