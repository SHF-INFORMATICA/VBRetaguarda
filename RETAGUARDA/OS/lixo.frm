VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTOTALSERVIÇO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtDESCONTOSERVIÇO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox txtTOTALPRODUTO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10545
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtDESCONTOPRODUTO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Serviços ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   7
      Top             =   6960
      Width           =   1710
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desc.Serviço ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   6
      Top             =   6960
      Width           =   1590
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Peças ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   5
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desc.Peças ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   4
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   6000
      X2              =   6000
      Y1              =   6840
      Y2              =   7320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim TOTAL_DESCONTO_SERVIÇO_N As Double, TOTAL_DESCONTO_PEÇAS_N As Double
   Dim IMPRESSORA As Printer, CONT_CURSO As Long
   Dim CONT_LINHAS As Long, PAGINA As Long
   Dim PaginaInicial, Paginafinal, NumeroDeCopias, i

Private Sub Form_Activate()

   txtDtIni.PromptInclude = False
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True

   If INDR_PRI = True Then
       cmbAUX.Clear
       cmbTipoOS.Clear
       SQL = "select * from DESCR "
       SQL = SQL & "where tipo_a = 'H' "
       SQL = SQL & "order by desc_a"
       Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
       While Not TabDESCR.EOF
          cmbTipoOS.AddItem Trim(TabDESCR!desc_a) & " - " & TabDESCR!Codigo
          cmbAUX.AddItem TabDESCR!Codigo
          TabDESCR.MoveNext
       Wend
       TabDESCR.Close
    
       cmbStatus.Clear
       cmbStatus.AddItem "A - ATIVA"
       'cmbSTATUS.AddItem "B - BAIXADA"
       'cmbSTATUS.AddItem "C - CANCELADA"
       cmbStatus.AddItem "D - NEGOCIAÇÂO"
       'cmbSTATUS.AddItem "E - EXECUSÃO"
       'cmbSTATUS.AddItem "F - FECHADA"
    
       cmbAuxMecanico.Clear
       cmbMecanico.Clear
       SQL = "select * from USUARIO "
       SQL = SQL & " where tipo = 8 "
       SQL = SQL & " and empresa_id = " & EMPRESA_ID
       SQL = SQL & " order by nome "
       Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
       While Not TabDESCR.EOF
          cmbMecanico.AddItem TabDESCR!NOME & " - " & TabDESCR!Codigo
          cmbAuxMecanico.AddItem TabDESCR!Codigo
          TabDESCR.MoveNext
       Wend
       TabDESCR.Close
    
       cmbVendedor.Clear
       cmbAuxVendedor.Clear
       SQL = "select * from VENDEDOR "
       SQL = SQL & "where status='A' " 'vendedores
       SQL = SQL & " order by nome_vend "
       Set TabDESCR = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
       While Not TabDESCR.EOF
          cmbVendedor.AddItem TabDESCR!NOME_VEND & "-" & TabDESCR!codg_vend
          cmbAuxVendedor.AddItem TabDESCR!codg_vend
          TabDESCR.MoveNext
       Wend
       TabDESCR.Close
      INDR_PRI = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape
         If txtOs.Text <> "" And INDR_GRAVA = True Then
            Msg = "Deseja sair sem gravar?"
            Style = vbYesNo + vbCritical
            Title = "Atenção !!!"
            Help = "DEMO.HLP"
            Ctxt = 1000
            RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
            If RESPOSTA = vbYes Then
               DBARQAUX.Execute "delete * from ITEMOS where numr_os =" & txtOs.Text
               Unload Me
            End If
            Else: Unload Me
         End If
      'Case vbKeyF6: EXCLUIR_ITEM
      Case vbKeyF9
         LIMPA_TUDO
         txtOs.SetFocus
      Case vbKeyF10
         If txtOs.Text <> "" Then _
            IMPRIMIR_OS
   End Select
End Sub

Private Sub Form_Load()
   Call CentralizaJanela(frmOSABRE)
   frmOSABRE.Top = 0
End Sub

Private Sub TXTCNPJCPF_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Clientes."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub TXTCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSCONSULTAVEICULO.Show 1
         If SQL3 <> "" Then _
            txtPlaca.Text = SQL3
         SQL3 = ""
         txtPlaca.SetFocus
   End Select
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         SQL = "select nome from CLIENTE "
         SQL = SQL & "where cgccpf = '" & txtCNPJCPF.Text & "'"
         Set tabcli = DBARQEMP.OpenRecordset(SQL, 4)
         If Not tabcli.EOF Then
            txtCli.Text = tabcli!NOME
            Else
               txtCNPJCPF.SelStart = 0
               txtCNPJCPF.SelLength = Len(txtOs)
               MsgBox "Cliente não Cadastrado."
               txtCNPJCPF.SetFocus
               Exit Sub
         End If
         Else
            txtPlaca.SetFocus
            Exit Sub
      End If
      cmbStatus.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub TXTCNPJCPF_LostFocus()
   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text <> "" Then
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
                  MsgBox "CNPJ com DV incorreto !!! "
                  txtCNPJCPF.PromptInclude = False
                  txtCNPJCPF = ""
                  txtCNPJCPF.SetFocus
                  Exit Sub
               End If
            Case Is > 14
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCNPJCPF = ""
               txtCNPJCPF.SetFocus
               Exit Sub
            Case Is < 11
               MsgBox "CNPJ/CPF com DV incorreto !!! "
               txtCNPJCPF = ""
               txtCNPJCPF.SetFocus
               Exit Sub
         End Select
         Else
            MsgBox "CNPJ/CPF com DV incorreto !!! "
            txtCNPJCPF = ""
            txtCNPJCPF.SetFocus
            Exit Sub
      End If
   End If
   txtCNPJCPF.PromptInclude = True
End Sub

Private Sub cmbVENDEDOR_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Vendedor"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   cmbVendedor.Text = "Oficina"
   cmbAuxVendedor.Text = "9999"
End Sub

Private Sub txtOs_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta O.S."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtOs_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         NUMR_OS = 0
         frmCONSULTAOS.Show 1
         If Not IsNull(NUMR_OS) Then _
            If NUMR_OS > 0 Then _
               txtOs.Text = NUMR_OS
   End Select
End Sub

Private Sub txtOS_KeyPress(KeyAscii As Integer)
   'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      NUMR_SEQ_N = 0
      If txtOs.Text = "" Then
         NUMR_OS = 0
         SQL = "select * from EMPRESA "
         SQL = SQL & "where empresa_id = " & EMPRESA_ID
         Set TabEMP = DBARQEMP.OpenRecordset(SQL)
         TabEMP.Edit
            TabEMP!seq_reqorc = TabEMP!seq_reqorc + 1
         TabEMP.Update
         NUMR_OS = TabEMP!seq_reqorc
         txtOs.Text = NUMR_OS
         TabEMP.Close
         Else: NUMR_OS = txtOs.Text
      End If

      ABRE_BANCO_AUXILIAR

      SQL = "select * from CABECAOS "
      SQL = SQL & "where numr_os = " & NUMR_OS
      Set TabCABECA = DBARQAUX.OpenRecordset(SQL, 4)
      If Not TabCABECA.EOF Then
         TRATA_OS
         If TabCABECA!Status = "F" Then
            txtOs.SelStart = 0
            txtOs.SelLength = Len(txtOs)
            txtOs.SetFocus
            MsgBox "O.S. Fechada.", vbOKOnly, "Atenção !!!"
            LIMPA_TUDO
            Exit Sub
         End If
         If TabCABECA!Status = "C" Then
            txtOs.SelStart = 0
            txtOs.SelLength = Len(txtOs)
            txtOs.SetFocus
            MsgBox "O.S. Cancelada.", vbOKOnly, "Atenção !!!"
            LIMPA_TUDO
            Exit Sub
         End If
         If TabCABECA!Status = "B" Then
            txtOs.SelStart = 0
            txtOs.SelLength = Len(txtOs)
            txtOs.SetFocus
            MsgBox "O.S. Baixada.", vbOKOnly, "Atenção !!!"
            LIMPA_TUDO
            Exit Sub
         End If
      End If
      TabCABECA.Close
      NUMR_OS = txtOs.Text
      txtDtIni.Text = Date
      txtCt.SetFocus
      'DBARQAUX.Close
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Me.Name, txtOs.Name
End Sub

Private Sub txtCli_gotfocus()
   'TXTCNPJCPF.SetFocus
End Sub

Private Sub txtCt_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta CT."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   LISTACT.Visible = True
   LISTACT.ListItems.Clear
   SQL = "select * from USUARIO "
   SQL = SQL & " where tipo = 6 "
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TabUSU = DBARQEMP.OpenRecordset(SQL, 4)
   While Not TabUSU.EOF
      Set Item = LISTACT.ListItems.Add(, "seq." & TabUSU!Codigo, TabUSU!Codigo)
      Item.SubItems(1) = TabUSU!NOME
      Item.SubItems(2) = TabUSU!Status
      TabUSU.MoveNext
   Wend
   TabUSU.Close
End Sub

Private Sub txtCT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtCt.Text = "" Then
         txtCt.Text = "9999"
         txtNomeCt.Text = "Consultor Geral"
         Else
            SQL = "select * from USUARIO "
            SQL = SQL & " where codigo = " & txtCt.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TabUSU = DBARQEMP.OpenRecordset(SQL, 4)
            If TabUSU.EOF Then
               MsgBox "Consultor não cadastrado Não Cadastrado.", vbOKOnly, "ERRO !!!"
               txtCt.SelStart = 0
               txtCt.SelLength = Len(txtCt)
               txtCt.SetFocus
               Exit Sub
               Else
                  If TabUSU!tipo = 6 Then
                     txtNomeCt.Text = TabUSU!NOME
                     'TXTCNPJCPF.SetFocus
                     Else
                        txtCt.SelStart = 0
                        txtCt.SelLength = Len(txtCt)
                        txtCt.SetFocus
                        MsgBox "Usuário não é Consultor Técnico."
                        txtCt.SetFocus
                        Exit Sub
                  End If
            End If
      End If
      cmbTipoOS.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtCt_LostFocus()
   LISTACT.Visible = False
End Sub

Private Sub txtDesc_Tarefa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtVALOR_TAREFA.SetFocus
   End If
End Sub

Private Sub txtDesc_TarefaS_GotFocus()
   cmbTipoOS.SetFocus
End Sub

Private Sub cmbmecanico_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione Mecânico"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
End Sub

Private Sub cmbmecanico_Click()
   cmbAuxMecanico.ListIndex = cmbMecanico.ListIndex
End Sub
   
Private Sub cmbmecanico_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbMecanico.Text = "" Then
         cmbMecanico.Text = "Oficina"
         cmbAuxMecanico.Text = "9999"
         Else
            SQL = "select * from USUARIO "
            SQL = SQL & " where codigo = " & cmbAuxMecanico.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TabDESCR.EOF Then
            If TabDESCR.EOF Then
               If TabDESCR!tipo <> 8 Then
                  TabDESCR.Close
                  MsgBox "Permitido somente mecanico."
                  Exit Sub
               End If
            End If
            End If
            TabDESCR.Close
      End If
      KeyAscii = 0
      txtDESCONTO_TAREFA.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtDESCONTO_TAREFA_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Valor de Desconto da tarefa"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtDESCONTO_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDESCONTO_TAREFA.Text <> "" Then
         If txtVALOR_TAREFA.Text <> "" Then
            VALOR_DESCONTO_N = txtDESCONTO_TAREFA.Text
            If VALOR_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_TAREFA.Text
               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               txtValor_Total_Tarefa.Text = Format(VALOR_TOTAL_N, "fixed")
               txtVALOR_TAREFA.Refresh

               txtPERC_TAREFA.Text = Format(((VALOR_DESCONTO_N * VALOR_ITEM_N) / 100), "fixed")
               txtPERC_TAREFA.Refresh
               GRAVA_ITEM_OS
               LIMPA_BODY_SERVIÇO
               Else: txtPERC_TAREFA.SetFocus
            End If
            Else: txtCODG_TAREFA.SetFocus
         End If
         Else: txtPERC_TAREFA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtPERC_TAREFA_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Percentual de Desconto da tarefa"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtPERC_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtVALOR_TAREFA.Text <> "" Then
         If txtPERC_TAREFA.Text <> "" Then
            PERC_DESCONTO_N = txtPERC_TAREFA.Text
            If PERC_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_TAREFA.Text

               VALOR_DESCONTO_N = ((PERC_DESCONTO_N * VALOR_ITEM_N) / 100)
               txtDESCONTO_TAREFA.Text = VALOR_DESCONTO_N

               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               txtValor_Total_Tarefa.Text = Format(VALOR_TOTAL_N, "fixed")
               txtValor_Total_Tarefa.Refresh

               GRAVA_ITEM_OS
               LIMPA_BODY_SERVIÇO
               Else
                  txtVALOR_TAREFA.Enabled = True
                  txtVALOR_TAREFA.SetFocus
            End If
            Else
               txtVALOR_TAREFA.Enabled = True
               txtVALOR_TAREFA.SetFocus
         End If
         Else
            txtVALOR_TAREFA.Enabled = True
            txtVALOR_TAREFA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtVALOR_TAREFA_GotFocus()
   txtVALOR_TAREFA.SelStart = 0
   txtVALOR_TAREFA.SelLength = Len(txtVALOR_TAREFA)
End Sub

Private Sub txtVALOR_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_ITEM_OS
      LIMPA_BODY_SERVIÇO
      txtCODG_TAREFA.SetFocus
   End If
End Sub

Private Sub cmbStatus_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione tipo de Ordem se Serviço"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbStatus.Text = "" Then _
         cmbStatus.Text = "A - Ativa"
      KeyAscii = 0
      txtKM.SetFocus
   End If
End Sub

Private Sub txtkm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbStatus.Text = "" Then _
         cmbStatus.Text = "A - Ativa"
      KeyAscii = 0
      txtCODG_TAREFA.SetFocus
   End If
End Sub

Private Sub txtCODG_TAREFA_GotFocus()
   txtVALOR_TAREFA.Enabled = False
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Tarefas."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F6 - Exclui Tarefas."
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (5)
   frmINICIO.BARI.Panels(5).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(5).AutoSize = sbrContents
   
End Sub

Private Sub txtCODG_TAREFA_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF6
         If Trim(txtCODG_TAREFA.Text) <> "" Then
            ABRE_BANCO_AUXILIAR
            SQL = "select * from ITEMOS "
            SQL = SQL & "where numr_os = " & NUMR_OS
            Set TabAUX = DBARQAUX.OpenRecordset(SQL)
            If Not TabAUX.EOF Then
               Msg = "Confirma exclusão dessa tarefa ?"
               PERGUNTA
               If RESPOSTA = vbYes Then
                  TabAUX.Delete
                  SETA_GRID_SERVIÇO
                  ATUALIZA_TOTAL_OS
                  txtCODG_TAREFA.SetFocus
               End If
            End If
            'TABAUX.Close
            'DBARQAUX.Close
         End If
      Case vbKeyF7
         CODG_PROD_A = ""
         frmCONSULTATAREFA.Show 1
         If CODG_PROD_A <> "" Then _
            txtCODG_TAREFA.Text = CODG_PROD_A
         CODG_PROD_A = ""
         txtCODG_TAREFA.SetFocus
   End Select
End Sub

Private Sub txtCODG_TAREFA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtCODG_TAREFA.Text) = "" Then
         MsgBox "Tarefa deve ser informada."
         Exit Sub
         Else
            ABRE_BANCO_AUXILIAR
            SQL = "select * from TAREFA "
            SQL = SQL & "where codg_tarefa = '" & Trim(txtCODG_TAREFA.Text) & "'"
            Set TabTemp = DBARQAUX.OpenRecordset(SQL, 4)
            If Not TabTemp.EOF Then
               txtDesc_Tarefa.Text = TabTemp!Descricao
               txtVALOR_TAREFA.Text = Format(TabTemp!valor_tarefa, "fixed")

               SQL = "select * from ITEMOS "
               SQL = SQL & "where numr_os = " & NUMR_OS
               SQL = SQL & " and codg_tarefa = '" & TabTemp!Codg_tarefa & "'"
               Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
               If Not TabAUX.EOF Then
                  txtDESCONTO_TAREFA.Text = Format(TabAUX!valor_desc_tarefa, "fixed")
                  txtVALOR_TAREFA.Text = Format(TabAUX!valor_tarefa, "fixed")
                  cmbAuxMecanico.Text = TabAUX!codg_mecanico

                  SQL = "select * from USUARIO "
                  SQL = SQL & " where codigo = " & TabAUX!codg_mecanico
                  SQL = SQL & " and empresa_id = " & EMPRESA_ID
                  Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
                  If Not TabDESCR.EOF Then _
                     cmbMecanico.Text = TabDESCR!NOME & " - " & TabDESCR!Codigo
                  TabDESCR.Close
               End If
               TabAUX.Close
               Else
                  TabTemp.Close
                  MsgBox "Tarefa não cadastrada, verifique."
                  Exit Sub
            End If
            TabTemp.Close
            DBARQAUX.Close
      End If
      cmbMecanico.SetFocus
   End If
End Sub

Private Sub cmbTipoOS_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Selecione Tipos O.S.."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub cmbtipoos_Click()
   cmbAUX.ListIndex = cmbTipoOS.ListIndex
End Sub
   
Private Sub cmbtipoos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbTipoOS.Text = "" Then
         cmbTipoOS.Text = "Normal"
         cmbAUX.Text = 1
      End If
      KeyAscii = 0
      txtPlaca.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtCHASSI_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Chassi"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtCHASSI_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPlaca.Text = "" Then
         'MsgBox "Chassi deve ser informado."
         'txtplaca.SetFocus
         'Exit Sub
         txtPlaca.SetFocus
         Exit Sub
         Else
            ABRE_BANCO_AUXILIAR
            PROCURA_PLACA
            DBARQAUX.Close
      End If
      cmbStatus.SetFocus
   End If
End Sub

Private Sub txtCHASSI_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSCONSULTAVEICULO.Show 1
         If SQL3 <> "" Then _
            txtPlaca.Text = SQL3
         SQL3 = ""
         txtPlaca.SetFocus
   End Select
End Sub
'==========
Private Sub txtPlaca_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F7 - Consulta Placa Veículo"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0
      If txtPlaca.Text = "" Then
         MsgBox "Placa deve ser informado."
         txtPlaca.SetFocus
         Exit Sub
         Else
            ABRE_BANCO_AUXILIAR
            SQL = "select placa from VEICULO "
            SQL = SQL & " where placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
            Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
            If TabAUX.EOF Then
               MsgBox "Placa não cadastrado."
               txtPlaca.SetFocus
               Exit Sub
               Else: PROCURA_PLACA
            End If
      End If
      cmbStatus.SetFocus
      Else
         If KeyAscii <> 8 Then
            CRITERIO = txtPlaca.Text
            If Len(CRITERIO) = 3 Then
               txtPlaca.Text = CRITERIO & "-"
               txtPlaca.SelStart = 4
               txtPlaca.Refresh
            End If
        End If
   End If
End Sub

Private Sub txtPLACA_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         SQL3 = ""
         frmOSCONSULTAVEICULO.Show 1
         If SQL3 <> "" Then
            ABRE_BANCO_AUXILIAR
            SQL = "select placa from VEICULO "
            SQL = SQL & "where chassi = '" & SQL3 & "'"
            Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
            If Not TabAUX.EOF Then
               txtPlaca.Text = Left(TabAUX!placa, 3) & "-" & Right(TabAUX!placa, 5)
            End If
            TabAUX.Close
            DBARQAUX.Close
         End If
         SQL3 = ""
         txtPlaca.SetFocus
   End Select
End Sub
'============================ PEÇAS
Private Sub txtproduto_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "F6 - Excluir Peça"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Produtos"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (5)
   frmINICIO.BARI.Panels(5).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(5).AutoSize = sbrContents
   
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF6
         If txtOs.Text <> "" And txtPRODUTO.Text <> "" Then _
            MATA_SEQ
      Case vbKeyF7
         frmCONSULTAPRODUTO.Show 1
         If CODG_PROD_A <> "" Then
            txtPRODUTO.Text = CODG_PROD_A
            txtPRODUTO.SetFocus
         End If
   End Select
End Sub

Private Sub txtproduto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      SQL = "select * from PRODUTO "
      SQL = SQL & " where codg_prod = '" & txtPRODUTO.Text & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID
      Set TabProduto = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
      If TabProduto.EOF Then
         'MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção !!!"
         'txtPRODUTO.SelStart = 0
         'txtPRODUTO.SelLength = Len(txtPRODUTO)
         'txtPRODUTO.SetFocus
         'Exit Sub
         cmbVendedor.SetFocus
         Exit Sub
         Else
            txtDESCPRODUTO.Text = TabProduto!Descricao
            'frmINICIO.BARI.Panels(3).Text = "Quantidade em Estoque = " & _
               TABPRODUTO!qtd - TABPRODUTO!qtd_balcao
            'frmINICIO.BARI.Panels(3).AutoSize = sbrContents

            QTD_ESTOQUE = TabProduto!QTD - TabProduto!qtd_balcao
            If Not IsNull(TabProduto!PRECO_VENDA) Then
               'frmINICIO.BARI.Panels(4).Text = Format(TABPRODUTO!PRECO_venda, "fixed")
               txtVALOR_PEÇA.Text = Format(TabProduto!PRECO_VENDA, "fixed")
            End If
            If txtOs.Text = "" Or txtPRODUTO.Text = "" Then _
               Exit Sub
            SQL = "select * from ITEMREQ "
            SQL = SQL & "where codg_prod = '" & txtPRODUTO.Text & "'"
            SQL = SQL & " and numr_req = " & txtOs.Text
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TABREQITEM = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
            If Not TABREQITEM.EOF Then
               'txtSeq.Text = tabreqitem!seq
               txtVALOR_PEÇA.Text = TABREQITEM!Valor_Item
               txtDESCONTO_PEÇA.Text = Format(TABREQITEM!PERC_DESCONTO, "fixed")
               txtQtd.Text = TABREQITEM!QTD_PEDIDA
               QTD_PEDIDO = TABREQITEM!SEQ
               QTD_EXTORNO_BALCAO = TABREQITEM!QTD_PEDIDA
               VALOR_ITEM_N = TABREQITEM!Valor_Item
               VALOR_DIFERENCA_N = TABREQITEM!VALOR_TOTAL_ITEM
               MsgBox "Produto já consta nesse O.S. seqüência = " & TABREQITEM!SEQ
               QTD_ESTOQUE = TabProduto!QTD + QTD_EXTORNO_BALCAO - TabProduto!qtd_balcao
            End If
            TabProduto.Close
            TABREQITEM.Close
      End If
      cmbVendedor.SetFocus
   End If
End Sub

Private Sub cmbVENDEDOR_Click()
   cmbAuxVendedor.ListIndex = cmbVendedor.ListIndex
End Sub
   
Private Sub cmbvendedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbVendedor.Text = "" Then
         cmbVendedor.Text = "Oficina"
         cmbAuxVendedor.Text = "9999"
         Else
            SQL = "select * from VENDEDOR "
            SQL = SQL & "where codg_vend = " & cmbAuxVendedor.Text
            Set TabDESCR = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
            If TabDESCR.EOF Then
               MsgBox "Vendedor não cadastrado."
               cmbVendedor.SetFocus
               Exit Sub
            End If
            TabDESCR.Close
      End If
      KeyAscii = 0
      txtQtd.SetFocus
      Else: KeyAscii = 0
   End If
End Sub

Private Sub txtQtd_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Quantidade"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Quantidade Disponível = " & QTD_ESTOQUE - QTD_PEDIDO
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (5)
   frmINICIO.BARI.Panels(5).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(5).AutoSize = sbrContents
   
   If txtPRODUTO.Text = Empty Then
      MsgBox "Codigo Produto inválido.", vbOKOnly, "Erro !!!"
      txtPRODUTO.Text = 99999999
      txtPRODUTO.SetFocus
      Exit Sub
   End If
   If txtQtd.Text <> "" Then
      txtQtd.SelStart = 0
      txtQtd.SelLength = Len(txtQtd)
   End If
End Sub

Private Sub txtqtd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbAuxVendedor.Text = "" Then
         Beep
         MsgBox "Selecione Vendedor.", vbOKOnly, "Atenção !!!"
         cmbVendedor.SetFocus
         Exit Sub
      End If
      If txtOs.Text = "" Then
         Beep
         MsgBox "Número de O.S. Inválido.", vbOKOnly, "Atenção !!!"
         txtOs.SetFocus
         Exit Sub
      End If
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text = "" Then
         Beep
         MsgBox "Informe Cliente.", vbOKOnly, "Atenção !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
      End If
      If txtPRODUTO.Text = "" Then
         MsgBox "Seqüência sem codigo de Produto.", vbOKOnly, "Atenção !!!"
         txtPRODUTO.SetFocus
         Exit Sub
      End If
      VALOR_ITEM_N = 0
      VALOR_ITEM_N = 0 & txtVALOR_PEÇA.Text
      If Not IsNull(VALOR_ITEM_N) Then
         If VALOR_ITEM_N <= 0 Then
            MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção !!!"
            txtPRODUTO.SetFocus
            Exit Sub
         End If
      End If
      KeyAscii = 0
      If txtQtd.Text = "" Then
         Beep
         MsgBox "Informe a quantidade.", vbOKOnly, "Atenção !!!"
         txtQtd.SetFocus
         Exit Sub
         Else
            'quantidade pedida
            QTD_PEDIDO = txtQtd.Text
            txtQtd.Text = Format(txtQtd.Text, "###.000")
            If INDR_CONTROLA_ESTOQUE = True Then
               If QTD_ESTOQUE < QTD_PEDIDO Then
                  Beep
                  MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção !!!"
                  txtQtd.SetFocus
                  Exit Sub
               End If
            End If
            If QTD_PEDIDO <= 0 Then
               Beep
               MsgBox "Quantidade pedida não permitido, deve ser maior que 0.", vbOKOnly, "Atenção !!!"
               txtQtd.SetFocus
               Exit Sub
            End If
      End If
      txtDESCONTO_PEÇA.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
End Sub

Private Sub txtDESCONTO_peça_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Valor de Desconto"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtDESCONTO_peça_keypress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDESCONTO_PEÇA.Text <> "" Then
         If txtVALOR_PEÇA.Text <> "" Then
            VALOR_DESCONTO_N = txtDESCONTO_PEÇA.Text
            If VALOR_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_PEÇA.Text
               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               'txtTOTALPRODUTO.Text = Format(VALOR_TOTAL_N, "fixed")
               'txtTOTALPRODUTO.Refresh

               txtPERC_PEÇA.Text = Format(((VALOR_DESCONTO_N * VALOR_ITEM_N) / 100), "fixed")
               txtPERC_PEÇA.Refresh
               GRAVA_CABECA
               LIMPA_BODY_PEÇA
               Else: txtPERC_PEÇA.SetFocus
            End If
            Else: txtPRODUTO.SetFocus
         End If
         Else: txtPERC_PEÇA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtPERC_PEÇA_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Percentual de Desconto"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
   
End Sub

Private Sub txtPERC_peça_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtVALOR_PEÇA.Text <> "" Then
         If txtPERC_PEÇA.Text <> "" Then
            PERC_DESCONTO_N = txtPERC_PEÇA.Text
            If PERC_DESCONTO_N > 0 Then
               VALOR_ITEM_N = txtVALOR_PEÇA.Text

               VALOR_DESCONTO_N = ((PERC_DESCONTO_N * VALOR_ITEM_N) / 100)
               txtDESCONTO_PEÇA.Text = VALOR_DESCONTO_N

               If VALOR_DESCONTO_N >= VALOR_ITEM_N Then
                  MsgBox "Valor de desconto inválido."
                  Exit Sub
               End If
               VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N
               'txtTOTALPRODUTO.Text = Format(VALOR_TOTAL_N, "fixed")
               'txtTOTALPRODUTO.Refresh

               GRAVA_CABECA
               LIMPA_BODY_PEÇA
               Else
                  txtVALOR_PEÇA.Enabled = True
                  txtVALOR_PEÇA.SetFocus
            End If
            Else
               txtVALOR_PEÇA.Enabled = True
               txtVALOR_PEÇA.SetFocus
         End If
         Else
            txtVALOR_PEÇA.Enabled = True
            txtVALOR_PEÇA.SetFocus
      End If
      KeyAscii = 0
   End If
End Sub

Private Sub txtVALOR_peça_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_CABECA
      LIMPA_BODY_PEÇA
   End If
End Sub

Private Sub txtvalor_peça_GotFocus()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Informe Valor Peça"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents


   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "F9 - Limpar "
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "F10 - Imprimir"
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents

   txtVALOR_PEÇA.SelStart = 0
   txtVALOR_PEÇA.SelLength = Len(txtVALOR_PEÇA)
End Sub

'=================================================================
Private Sub GRAVA_CABECA()
   ABRE_BANCO_AUXILIAR
   GRAVA_CABECA_OS
   SQL = "select * from CABECAREQ "
   SQL = SQL & " where numr_req = " & txtOs.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TabCABECA = DBARQEMP.OpenRecordset(SQL)
   If TabCABECA.EOF Then
      TabCABECA.AddNew
      Else: TabCABECA.Edit
   End If
   GRAVA_PEÇA
   TabCABECA.Close
End Sub

Private Sub GRAVA_PEÇA()
   txtCNPJCPF.PromptInclude = False
   SQL = "select cgccpf from CLIENTE "
   SQL = SQL & "where cgccpf = '" & txtCNPJCPF.Text & "'"
   Set tabcli = DBARQEMP.OpenRecordset(SQL, 4)
   If tabcli.EOF Then
      tabcli.Close
      MsgBox "Cliente não cadastrado, verifique."
      Exit Sub
   End If
   tabcli.Close
   TabCABECA!numr_req = txtOs.Text
   txtCNPJCPF.PromptInclude = False
   TabCABECA!CGCCPF = txtCNPJCPF.Text
   If cmbAuxVendedor.Text = "" Then
      cmbVendedor.Text = "Oficina"
      cmbAuxVendedor.Text = "9999"
   End If
   TabCABECA!Vendedor = cmbAuxVendedor.Text
   TabCABECA!DT_REQ = Date
   TabCABECA!Valor_Total = VALOR_TOTAL_N
   TabCABECA!tipovenda_id = 1
   'AGORA TODAS VENDAS a vista vai para emitir cupom ou nota
   TabCABECA!Status = 2
   If txtDESCONTO_PEÇA.Text = "" Then
      TabCABECA!Valor_Desconto = 0
      TabCABECA!PERC_desc = 0
      Else
         If txtDESCONTOPRODUTO.Text <> "" Then
            VALOR_TOTAL_DESCONTO_N = txtDESCONTOPRODUTO.Text
            TabCABECA!Valor_Desconto = txtDESCONTOPRODUTO.Text
         End If
         TabCABECA!PERC_desc = (VALOR_TOTAL_DESCONTO_N / 100)
   End If
   TabCABECA!TIPO_REGISTRO = "S"
   TabCABECA!USUARIO_LIBERA_VENDA = USUARIO_LIBERA_VENDA
   If Not IsNull(CODG_USU_N) Then _
      TabCABECA!CODG_USU = CODG_USU_N
   TabCABECA!EMPRESA_ID = EMPRESA_ID
   TabCABECA.Update
   GRAVA_PEÇA_ITEM
End Sub

Private Sub GRAVA_PEÇA_ITEM()
   QTD_PEDIDO = txtQtd.Text
   VALOR_ITEM_N = txtVALOR_PEÇA.Text
   PERC_DESCONTO_N = 0
   If txtPERC_PEÇA.Text <> "" Then _
      PERC_DESCONTO_N = txtPERC_PEÇA.Text
   If txtDESCONTO_PEÇA.Text <> "" Then _
      VALOR_DESCONTO_N = txtDESCONTO_PEÇA.Text
   NUMR_SEQ_N = 1
   SQL = "select max(seq) from ITEMREQ "
   SQL = SQL & "where numr_req = " & txtOs.Text
   Set TABREQITEM = DBARQEMP.OpenRecordset(SQL)
   If Not TABREQITEM.EOF Then _
      If Not IsNull(TABREQITEM.Fields(0).Value) Then _
         NUMR_SEQ_N = 1 + TABREQITEM.Fields(0).Value
   TABREQITEM.Close

   SQL = "select * from ITEMREQ "
   SQL = SQL & " where codg_prod = '" & txtPRODUTO.Text & "'"
   SQL = SQL & " and numr_req = " & txtOs.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TABREQITEM = DBARQEMP.OpenRecordset(SQL)
   If TABREQITEM.EOF Then
      TABREQITEM.AddNew
        TABREQITEM!numr_req = txtOs.Text
        TABREQITEM!SEQ = NUMR_SEQ_N
        TABREQITEM!Codg_Prod = txtPRODUTO.Text
      Else: TABREQITEM.Edit
   End If
   TABREQITEM!QTD_PEDIDA = QTD_PEDIDO
   TABREQITEM!Valor_Item = VALOR_ITEM_N
   TABREQITEM!PERC_desc = PERC_DESCONTO_N
   TABREQITEM!EMPRESA_ID = EMPRESA_ID
   'TABREQITEM!VALOR_TOTAL_ITEM = (VALOR_ITEM_N * QTD_PEDIDO) - VALOR_DESCONTO_N
   TABREQITEM.Update
   TABREQITEM.Close
   SETA_GRID_PEÇA
   ATUALIZA_TOTAL_OS
End Sub

Private Sub MATA_SEQ()
   SQL = "select * from ITEMREQ "
   SQL = SQL & " where codg_prod = '" & txtPRODUTO.Text & "'"
   SQL = SQL & " and numr_req = " & txtOs.Text
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TabTemp = DBARQEMP.OpenRecordset(SQL)
   If Not TabTemp.EOF Then
      Msg = "Deseja Excluir Esse Item?"
      Style = vbYesNo + 32
      Title = "Atenção !!!"
      Help = "DEMO.HLP"
      Ctxt = 1000
      RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
      If RESPOSTA = vbYes Then
         If INDR_ATUALIZA_ESTOQUE = False Then
            SQL = "select * from PRODUTO "
            SQL = SQL & " where codg_prod = '" & TabTemp!Codg_Prod & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TabProduto = DBARQEMP.OpenRecordset(SQL)
            If Not TabProduto.EOF Then
               TabProduto.Edit
                  TabProduto!qtd_balcao = TabProduto!qtd_balcao - TabTemp!QTD_PEDIDA
               TabProduto.Update
            End If
            TabProduto.Close
         End If
         TabTemp.Delete
         TabTemp.Close
         LIMPA_BODY_PEÇA
         SETA_GRID_PEÇA
         ATUALIZA_TOTAL_OS
         Else: TabTemp.Close
      End If
   End If
   txtPRODUTO.SetFocus
End Sub

Private Sub LIMPA_TUDO()
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   TOTAL_SERVIÇO_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0

   txtPlaca.Text = ""
   LISTASERVIÇO.ListItems.Clear
   txtKM.Text = ""
   LISTAPEÇA.ListItems.Clear
   TOTAL_SERVIÇO_N = 0
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   txtOs.Text = ""
   txtCt.Text = ""
   txtNomeCt.Text = ""
   cmbAUX.Text = ""
   cmbTipoOS.Text = ""
   txtDtIni.PromptInclude = False
      txtDtIni.Text = Date
   txtDtIni.PromptInclude = True
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtCli.Text = ""
   cmbStatus.Text = ""
   txtTOTALDESCONTOOS.Text = ""
   txtTOTALOS.Text = ""
   txtDESCONTOSERVIÇO.Text = ""
   txtTOTALSERVIÇO.Text = ""
   txtDESCONTOPRODUTO.Text = ""
   txtTOTALPRODUTO.Text = ""
   LIMPA_BODY_SERVIÇO
   LIMPA_BODY_PEÇA
End Sub

Private Sub LIMPA_BODY_SERVIÇO()
   txtCODG_TAREFA.Text = ""
   txtDesc_Tarefa.Text = ""
   txtDESCONTO_TAREFA.Text = ""
   txtPERC_TAREFA.Text = ""
   txtVALOR_TAREFA.Text = ""
   txtValor_Total_Tarefa.Text = ""
End Sub

Private Sub LIMPA_BODY_PEÇA()
   txtPRODUTO.Text = ""
   txtDESCPRODUTO.Text = ""
   cmbVendedor.Text = ""
   cmbAuxVendedor.Text = ""
   txtQtd.Text = ""
   txtDESCONTO_PEÇA.Text = ""
   txtPERC_PEÇA.Text = ""
   txtVALOR_PEÇA.Text = ""
   txtTOTAL_PEÇA.Text = ""
   txtPRODUTO.SetFocus
End Sub

Private Sub TRATA_OS()
   MOSTRA_OS
   SETA_GRID_SERVIÇO
   SETA_GRID_PEÇA
   ATUALIZA_TOTAL_OS
End Sub

Private Sub MOSTRA_OS()
   txtOs.Text = TabCABECA!NUMR_OS
   txtCt.Text = TabCABECA!ct
   If Not IsNull(TabCABECA!km_atual) Then
      txtKM.Text = TabCABECA!km_atual
      Else: txtKM.Text = ""
   End If
   SQL = "select nome from USUARIO "
   SQL = SQL & " where codigo = " & TabCABECA!ct
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TabUSU = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TabUSU.EOF Then _
      txtNomeCt.Text = TabUSU!NOME
   TabUSU.Close

   SQL = "select * from DESCR "
   SQL = SQL & "where tipo_a = 'H' "
   SQL = SQL & "and codigo = " & TabCABECA!TIPO_OS
   Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TabDESCR.EOF Then
      cmbTipoOS.Text = Trim(TabDESCR!desc_a) & " - " & TabDESCR!Codigo
      cmbAUX.Text = TabDESCR!Codigo
   End If
   TabDESCR.Close

   txtDtIni.PromptInclude = False
      txtDtIni.Text = TabCABECA!dt_abertura
   txtDtIni.PromptInclude = True

   SQL = "select * from VEICULO "
   SQL = SQL & "where placa = '" & Trim(TabCABECA!placa) & "'"
   Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
   If Not TabAUX.EOF Then
      txtPlaca.Text = TabAUX!placa
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = TabAUX!CGCCPF
      SQL = "select nome from CLIENTE "
      SQL = SQL & "where cgccpf = '" & TabAUX!CGCCPF & "'"
      Set tabcli = DBARQEMP.OpenRecordset(SQL, 4)
      If Not tabcli.EOF Then
         txtCli.Text = tabcli!NOME
         Else: MsgBox "Cliente não cadastrado, verifique."
      End If
   End If
   TabAUX.Close

   If TabCABECA!Status = "A" Then _
      cmbStatus.Text = "Aberta"
   If TabCABECA!Status = "B" Then _
      cmbStatus.Text = "Baixada"
   If TabCABECA!Status = "C" Then _
      cmbStatus.Text = "Cancelada"
   If TabCABECA!Status = "D" Then _
      cmbStatus.Text = "Em Negociação"
   If TabCABECA!Status = "E" Then _
      cmbStatus.Text = "Em Execução"
   If TabCABECA!Status = "F" Then _
      cmbStatus.Text = "Fechada"
End Sub

Private Sub SETA_GRID_SERVIÇO()
   LISTASERVIÇO.ListItems.Clear
   VALOR_TOTAL_N = 0
   TOTAL_SERVIÇO_N = 0
   VALOR_DESCONTO_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0
   NUMR_SEQ_N = 1
   SQL = "select * from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   SQL = SQL & " order by hora_inicio"
   Set TabTemp = DBARQAUX.OpenRecordset(SQL, 4)
   While Not TabTemp.EOF
      NUMR_SEQ_N = 1 + NUMR_SEQ_N
      Set Item = LISTASERVIÇO.ListItems.Add(, "seq." & NUMR_SEQ_N, TabTemp!Codg_tarefa)
      SQL = "select * from TAREFA "
      SQL = SQL & "where codg_tarefa = '" & TabTemp!Codg_tarefa & "'"
      Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
      If Not TabAUX.EOF Then _
         Item.SubItems(1) = TabAUX!Descricao
      TabAUX.Close
      TOTAL_SERVIÇO_N = TOTAL_SERVIÇO_N + TabTemp!valor_tarefa
      TOTAL_DESCONTO_SERVIÇO_N = TOTAL_DESCONTO_SERVIÇO_N + TabTemp!valor_desc_tarefa

      Item.SubItems(2) = Format(TabTemp!valor_tarefa, "fixed")
      Item.SubItems(3) = Format(TabTemp!valor_desc_tarefa, "fixed")
      Item.SubItems(4) = Format(TabTemp!valor_tarefa - TabTemp!valor_desc_tarefa, "fixed")
      If TabTemp!Status = "A" Then _
         Item.SubItems(5) = "Ativo"
      If TabTemp!Status = "B" Then _
         Item.SubItems(5) = "Baixado"
      If TabTemp!Status = "C" Then _
         Item.SubItems(5) = "Cancelado"
      If TabTemp!Status = "E" Then _
         Item.SubItems(5) = "Execução"
      SQL = "select * from USUARIO "
      SQL = SQL & "where codigo = " & TabTemp!codg_mecanico
      Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
      If Not TabDESCR.EOF Then _
         Item.SubItems(6) = TabDESCR!NOME & " - " & TabDESCR!Codigo
      TabDESCR.Close
      TabTemp.MoveNext
   Wend
   TabTemp.Close
   txtTOTALSERVIÇO.Text = Format(TOTAL_SERVIÇO_N - TOTAL_DESCONTO_SERVIÇO_N, "fixed")
   txtTOTALSERVIÇO.Refresh
   txtDESCONTOSERVIÇO.Text = Format(TOTAL_DESCONTO_SERVIÇO_N, "fixed")
   txtDESCONTOSERVIÇO.Refresh
End Sub

Private Sub SETA_GRID_PEÇA()
   LISTAPEÇA.ListItems.Clear
   NUMR_SEQ_N = 0
   VALOR_DESCONTO_N = 0
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   SQL = "select * from ITEMREQ "
   SQL = SQL & "where numr_req = " & txtOs.Text
   If NUMR_SEQ_N < 10 Then
      SQL = SQL & " order by seq asc"
      Else: SQL = SQL & " order by seq desc"
   End If
   Set TABREQITEM = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
   While Not TABREQITEM.EOF
      TOTAL_PEÇAS_N = TOTAL_PEÇAS_N + (TABREQITEM!Valor_Item * TABREQITEM!QTD_PEDIDA)
      TOTAL_DESCONTO_PEÇAS_N = TOTAL_DESCONTO_PEÇAS_N + (TABREQITEM!PERC_desc * TABREQITEM!Valor_Item / 100)

      Set Item = LISTAPEÇA.ListItems.Add(, "seq." & TABREQITEM!Codg_Prod, TABREQITEM!Codg_Prod)
      SQL = "select descricao,referencia from PRODUTO "
      SQL = SQL & "where codg_prod = '" & TABREQITEM!Codg_Prod & "'"
      Set TabTemp = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
      If Not TabTemp.EOF Then
         Item.SubItems(1) = TabTemp!Descricao
         If Not IsNull(TabTemp!Referencia) Then
            Item.SubItems(6) = TabTemp!Referencia
         End If
      End If
      TabTemp.Close
      Item.SubItems(2) = TABREQITEM!QTD_PEDIDA
      Item.SubItems(3) = Format(TABREQITEM!Valor_Item, "fixed")
      Item.SubItems(4) = Format(TABREQITEM!Valor_Item * TABREQITEM!PERC_desc / 100, "fixed")
      Item.SubItems(5) = Format(TABREQITEM!Valor_Item * TABREQITEM!QTD_PEDIDA - (TABREQITEM!Valor_Item * TABREQITEM!PERC_desc / 100), "fixed")
      TABREQITEM.MoveNext
   Wend
   TABREQITEM.Close
   txtTOTALPRODUTO.Text = Format(TOTAL_PEÇAS_N - TOTAL_DESCONTO_PEÇAS_N, "fixed")
   txtTOTALPRODUTO.Refresh
   txtDESCONTOPRODUTO.Text = Format(TOTAL_DESCONTO_PEÇAS_N, "fixed")
   txtDESCONTOPRODUTO.Refresh
End Sub

Private Sub GRAVA_CABECA_OS()
   SQL = "select * from CABECAOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TabCABECA = DBARQAUX.OpenRecordset(SQL)
   If Not TabCABECA.EOF Then
      TabCABECA.Edit
      Else
         TabCABECA.AddNew
            TabCABECA!dt_abertura = Now
   End If
   TabCABECA!NUMR_OS = NUMR_OS
   TabCABECA!placa = Replace(txtPlaca.Text, "-", "")
   If cmbAUX.Text <> "" Then
      TabCABECA!TIPO_OS = cmbAUX.Text
   End If
   If txtCt.Text <> "" Then
      TabCABECA!ct = txtCt.Text
   End If
   If cmbStatus.Text <> "" Then
      TabCABECA!Status = Left(cmbStatus.Text, 1)
   End If
   TabCABECA!km_atual = txtKM.Text
   'If txtTOTALDESCONTOOS.Text = "" Then
   '   TABCABECA!valor_desconto = 0
   '   Else: TABCABECA!valor_desconto = txtTOTALDESCONTOOS.Text
   'End If
   TabCABECA.Update
   TabCABECA.Close
End Sub

Private Sub GRAVA_ITEM_OS()
   If txtOs.Text = "" Then
      MsgBox "Informe número de O.S."
      txtOs.SetFocus
      Exit Sub
   End If
   If txtCt.Text = "" Then
      MsgBox "Informe Consultor Técnico"
      txtCt.SetFocus
      Exit Sub
   End If
   If cmbTipoOS.Text = "" Then
      MsgBox "Informe Tipo de O.S."
      cmbTipoOS.SetFocus
      Exit Sub
   End If
   If cmbStatus.Text = "" Then
      MsgBox "Informe status da O.S."
      cmbStatus.SetFocus
      Exit Sub
   End If
   If Trim(txtCODG_TAREFA.Text) = "" Then
      MsgBox "Informe Código da tarefa da O.S."
      txtCODG_TAREFA.SetFocus
      Exit Sub
   End If
   If cmbMecanico.Text = "" Then
      MsgBox "Informe mecanico dessa tarefa da O.S."
      cmbMecanico.SetFocus
      Exit Sub
   End If
   If txtVALOR_TAREFA.Text = "" Then
      MsgBox "Informe valor da tarefa dessa O.S."
      txtVALOR_TAREFA.SetFocus
      Exit Sub
   End If

   ABRE_BANCO_AUXILIAR

   GRAVA_CABECA_OS

   SQL = "select * from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   SQL = SQL & " and codg_tarefa = '" & Trim(txtCODG_TAREFA.Text) & "'"
   Set TabAUX = DBARQAUX.OpenRecordset(SQL)
   If TabAUX.EOF Then
      TabAUX.AddNew
         TabAUX!HORA_INICIO = Now
      Else: TabAUX.Edit
   End If
   TabAUX!NUMR_OS = NUMR_OS
   TabAUX!Codg_tarefa = Trim(txtCODG_TAREFA.Text)
   TabAUX!valor_tarefa = txtVALOR_TAREFA.Text
   If txtDESCONTO_TAREFA.Text <> "" Then
      TabAUX!valor_desc_tarefa = txtDESCONTO_TAREFA.Text
      Else: TabAUX!valor_desc_tarefa = 0
   End If
   TabAUX!Status = "A"
   TabAUX!codg_mecanico = cmbAuxMecanico.Text
   TabAUX.Update
   TabAUX.Close
   SETA_GRID_SERVIÇO
   txtCODG_TAREFA.SetFocus
   DBARQAUX.Close

   ATUALIZA_TOTAL_OS
End Sub

Private Sub PROCURA_PLACA()
   SQL = "select * from VEICULO "
   If txtPlaca.Text <> "" Then _
      SQL = SQL & " where placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
   Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
   If TabAUX.EOF Then
      MsgBox "Placa não cadastrado."
      txtPlaca.SetFocus
      Exit Sub
      Else
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = TabAUX!CGCCPF
         'txtPLACA.Text = Left(TABAUX!placa, 3) & "-" & Right(TABAUX!placa, 5)

         SQL = "select nome from CLIENTE "
         SQL = SQL & "where cgccpf = '" & txtCNPJCPF.Text & "'"
         Set tabcli = DBARQEMP.OpenRecordset(SQL, 4)
         If Not tabcli.EOF Then
            txtCli.Text = tabcli!NOME
            txtCNPJCPF.PromptInclude = False
            If txtCNPJCPF.Text <> "" Then
               If Len(txtCNPJCPF.Text) > 0 Then
                  Select Case Len(txtCNPJCPF.Text)
                     Case Is = 11
                        If Not CALCULACPF(txtCNPJCPF.Text) Then
                           MsgBox "CPF com DV incorreto !!!"
                           txtCNPJCPF.PromptInclude = False
                           'TXTCNPJCPF = ""
                           'TXTCNPJCPF.SetFocus
                           Exit Sub
                        End If
                     Case Is = 14
                        If Not VALIDACGC(txtCNPJCPF.Text) Then
                           MsgBox "CNPJ com DV incorreto !!! "
                           txtCNPJCPF.PromptInclude = False
                           'TXTCNPJCPF = ""
                           'TXTCNPJCPF.SetFocus
                           Exit Sub
                        End If
                     Case Is > 14
                        MsgBox "CNPJ/CPF com DV incorreto !!! "
                        'TXTCNPJCPF = ""
                        'TXTCNPJCPF.SetFocus
                        Exit Sub
                     Case Is < 11
                        MsgBox "CNPJ/CPF com DV incorreto !!! "
                        'TXTCNPJCPF = ""
                        'TXTCNPJCPF.SetFocus
                        Exit Sub
                  End Select
                  Else
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     'TXTCNPJCPF = ""
                     'TXTCNPJCPF.SetFocus
                     Exit Sub
               End If
            End If
            txtCNPJCPF.PromptInclude = True
         End If
   End If
   TabAUX.Close
End Sub

Private Sub ATUALIZA_TOTAL_OS()
   TOTAL_PEÇAS_N = 0
   TOTAL_DESCONTO_PEÇAS_N = 0
   TOTAL_SERVIÇO_N = 0
   TOTAL_DESCONTO_SERVIÇO_N = 0

   SQL = "select sum(valor_item*qtd_pedida) from ITEMREQ "
   SQL = SQL & "where numr_req = " & NUMR_OS
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TabConsulta = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         TOTAL_PEÇAS_N = TabConsulta.Fields(0).Value
   TabConsulta.Close

   SQL = "select sum(perc_desc) from ITEMREQ "
   SQL = SQL & " where numr_req = " & NUMR_OS
   SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TabConsulta = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         TOTAL_DESCONTO_PEÇAS_N = TOTAL_PEÇAS_N * TabConsulta.Fields(0).Value / 100
   TabConsulta.Close

   ABRE_BANCO_AUXILIAR

   SQL = "select sum(valor_tarefa) from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TabConsulta = DBARQAUX.OpenRecordset(SQL, 4)
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         TOTAL_SERVIÇO_N = TabConsulta.Fields(0).Value
   TabConsulta.Close

   SQL = "select sum(valor_desc_tarefa) from ITEMOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TabConsulta = DBARQAUX.OpenRecordset(SQL, 4)
   If Not TabConsulta.EOF Then _
      If Not IsNull(TabConsulta.Fields(0).Value) Then _
         TOTAL_DESCONTO_SERVIÇO_N = TabConsulta.Fields(0).Value
   TabConsulta.Close

   VALOR_DESCONTO_N = 0
   SQL = "select valor_desconto from CABECAOS "
   SQL = SQL & "where numr_os = " & NUMR_OS
   Set TabUSU = DBARQAUX.OpenRecordset(SQL, 4)
   If Not IsNull(TabUSU!Valor_Desconto) Then _
      VALOR_DESCONTO_N = TabUSU!Valor_Desconto

   txtTOTALOS.Text = Format((TOTAL_SERVIÇO_N + TOTAL_PEÇAS_N) - (TOTAL_DESCONTO_PEÇAS_N + TOTAL_DESCONTO_SERVIÇO_N + VALOR_DESCONTO_N), "fixed")
   txtTOTALOS.Refresh
   txtTOTALDESCONTOOS.Text = Format((TOTAL_DESCONTO_PEÇAS_N + TOTAL_DESCONTO_SERVIÇO_N + VALOR_DESCONTO_N), "fixed")
   txtTOTALDESCONTOOS.Refresh

   DBARQAUX.Close
End Sub
'================
Private Sub IMPRIMIR_OS()
   If txtOs.Text = "" Then _
      Exit Sub
   OBS_A = ""
   If txtTOTALOS.Text <> "" Then
      VALOR_TOTAL_N = txtTOTALOS.Text
      VALOR_ITEM_N = txtTOTALOS.Text
   End If
   VALOR_TOTAL_DESCONTO_N = 0
   frmDESCONTO.Show 1

   ABRE_BANCO_AUXILIAR
   SQL = "update CABECAOS set "
   SQL = SQL & " valor_desconto = '" & VALOR_TOTAL_DESCONTO_N & "'"
   SQL = SQL & " where numr_os = " & NUMR_OS
   'SQL = SQL & " and empresa_id = " & EMPRESA_ID
   DBARQAUX.Execute SQL

   SQL = "select * from CABECAOS "
   SQL = SQL & " where numr_os = " & NUMR_OS
   'SQL = SQL & " and empresa_id = " & EMPRESA_ID
   Set TabCABECA = DBARQAUX.OpenRecordset(SQL)
   If Not TabCABECA.EOF Then
      TabCABECA.Edit
         TabCABECA!Valor_Desconto = VALOR_TOTAL_DESCONTO_N
         'TABCABECA!EMPRESA_ID = EMPRESA_ID
      TabCABECA.Update
   End If
   TabCABECA.Close

   SQL = "update RELOS set "
   SQL = SQL & "valor_desconto = '" & Replace(VALOR_TOTAL_DESCONTO_N, ",", ".") & "'"
   SQL = SQL & ", obs = '" & OBS_A & "'"
   SQL = SQL & " where numr_os = " & NUMR_OS
   DBARQAUX.Execute SQL

   If txtDESCONTOPRODUTO.Text <> "" Or txtTOTALPRODUTO.Text <> "" Then
      SQL = "update CABECAREQ set "
      'SQL = SQL & "valor_desconto = '" & Replace(txtDESCONTOPRODUTO.Text, ",", ".") & "'"
      SQL = SQL & " valor_total = '" & Replace(txtTOTALPRODUTO.Text, ",", ".") & "'"
      SQL = SQL & " where numr_req = " & NUMR_OS
      SQL = SQL & " and empresa_id = " & EMPRESA_ID
      DBARQEMP.Execute SQL
   End If

   SQL = "select * from RELOS "
   SQL = SQL & " where numr_os = " & NUMR_OS
   Set TabCABECA = DBARQAUX.OpenRecordset(SQL)
   If Not TabCABECA.EOF Then
      TabCABECA.Edit
         TabCABECA!Valor_Desconto = VALOR_TOTAL_DESCONTO_N
      TabCABECA.Update
   End If
   TabCABECA.Close

   SQL = "update OBS set "
   SQL = SQL & " obs = '" & OBS_A & "'"
   SQL = SQL & " ,prop = " & NUMR_OS
   SQL = SQL & " ,seq = 1 "
   SQL = SQL & " where prop = " & NUMR_OS
   SQL = SQL & " and seq = 1 "
   DBARQEMP.Execute SQL

   SQL = "select * from OBS "
   SQL = SQL & " where prop = " & NUMR_OS
   Set TabCABECA = DBARQEMP.OpenRecordset(SQL)
   If Not TabCABECA.EOF Then
      TabCABECA.Edit
         TabCABECA!obs = OBS_A
         TabCABECA!prop = NUMR_OS
      TabCABECA.Update
   End If
   TabCABECA.Close

   SQL = "select cgc,razao_social,nome_fant,ie from EMPRESA "
   SQL = SQL & "where empresa_id = " & EMPRESA_ID
   Set TabEMP = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TabEMP.EOF Then
      txtCGC.PromptInclude = False
      txtCNPJCPF.PromptInclude = False
         txtCGC.Text = TabEMP!CGC
      txtCGC.PromptInclude = True

      ABRE_BANCO_AUXILIAR
      
      SQL = "create table RELOS "
      SQL = SQL & "("
         SQL = SQL & "cgc text(20), razao_social text(80),nome_fant text(80),ie "
         SQL = SQL & "text(40),end_emp text(40),bairro_emp text(40),cep_emp text(10),"
         SQL = SQL & "cidade_uf_emp text(40), nome_cli text(60),end_cli text(80),"
         SQL = SQL & "descricao_veiculo text(40), cor text(10),ano_modelo text(10),motor text(10),"
         SQL = SQL & "chassi text(80),combustivel text(10),placa text(10),dt_abre text(20),"
         SQL = SQL & "km text(10),dt_fecha text(20),consultor text(30),obs  text(255), tipo_os text(30),"
         SQL = SQL & "valor_desconto double,numr_os long not null,status text(1), fone text(50)"
         'SQL = SQL & " constraint numr_os unique (CHAVE_numr_os)"
      SQL = SQL & ")"
'MsgBox SQL
      'DBARQAUX.Execute SQL

      SQL = "create table RELOSITEM "
      SQL = SQL & "("
         SQL = SQL & "numr_os long not null,codg_item text(10),desc_item text(50),valor_item double"
         SQL = SQL & ",desconto_item double,qtd long,tipo_item text(1)"
      SQL = SQL & ")"
      'DBARQAUX.Execute SQL

      SQL = "delete * from RELOS "
      SQL = SQL & "where numr_os = " & txtOs.Text
      DBARQAUX.Execute SQL

      GRAVA_CABECA_OS

      SQL = "select * from CABECAOS "
      SQL = SQL & "where numr_os = " & txtOs.Text
      Set TabTemp = DBARQAUX.OpenRecordset(SQL, 4)
      If Not TabTemp.EOF Then
         SQL = "select * from RELOS "
         SQL = SQL & "where numr_os = " & TabTemp!NUMR_OS
         Set TabAUX = DBARQAUX.OpenRecordset(SQL)
         If Not TabAUX.EOF Then
            TabAUX.Edit
            Else: TabAUX.AddNew
         End If
         TabAUX!CGC = txtCGC.Text
         TabAUX!RAZAO_SOCIAL = TabEMP!RAZAO_SOCIAL
         TabAUX!Nome_Fant = TabEMP!Nome_Fant
         TabAUX!IE = TabEMP!IE
         TabAUX!dt_abre = TabTemp!dt_abertura
         TabAUX!dt_fecha = TabTemp!dt_fechamento
         TabAUX!Valor_Desconto = TabTemp!Valor_Desconto
         TabAUX!NUMR_OS = TabTemp!NUMR_OS
         TabAUX!Status = TabTemp!Status
         'TIPO OS
         SQL = "select * from DESCR "
         SQL = SQL & "where tipo_a = 'H' "
         SQL = SQL & "and codigo = " & TabTemp!TIPO_OS
         Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TabDESCR.EOF Then _
            TabAUX!TIPO_OS = Trim(TabDESCR!desc_a)
         TabDESCR.Close

         SQL = "select nome from USUARIO "
         SQL = SQL & "where codigo = " & TabTemp!ct
         Set TabUSU = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TabUSU.EOF Then _
            TabAUX!consultor = TabUSU!NOME
         TabUSU.Close

         TabAUX!obs = OBS_A

         CRITERIO = ""
         SQL = "select * from FONE "
         SQL = SQL & "where prop = '" & TabEMP!CGC & "'"
         Set TabEND = DBARQEMP.OpenRecordset(SQL, 4)
         While Not TabEND.EOF
            If Not IsNull(TabEND!DDD) Then _
               CRITERIO = CRITERIO & "   (" & TabEND!DDD & ") "
            If Not IsNull(TabEND!Numero) Then _
               CRITERIO = CRITERIO & TabEND!Numero
            TabEND.MoveNext
         Wend
         TabEND.Close
         TabAUX!FONE = Trim(CRITERIO)

         'ENDEREÇO EMPRESA
         SQL = "select * from ENDERECO "
         SQL = SQL & "where prop = '" & TabEMP!CGC & "'"
         SQL = SQL & " and tipo = 'C' "
         Set TabEND = DBARQEMP.OpenRecordset(SQL, 4)
         If Not TabEND.EOF Then
            If Not IsNull(TabEND!Rua) Then _
               TabAUX!end_emp = TabEND!Rua
            If Not IsNull(TabAUX!end_emp) Then
               If Not IsNull(TabEND!Complemento) Then _
                  TabAUX!end_emp = TabAUX!end_emp & " ; " & TabEND!Complemento
               Else
                  If Not IsNull(TabEND!Complemento) Then _
                     TabAUX!end_emp = TabEND!Complemento
            End If
            If Not IsNull(TabEND!Bairro) Then _
               TabAUX!bairro_emp = TabEND!Bairro
            If Not IsNull(TabEND!CEP) Then _
               TabAUX!cep_emp = TabEND!CEP
            If Not IsNull(TabEND!CEP) Then
               If TabEND!CEP <> "" Then
                  SQL = "select * from CEP "
                  SQL = SQL & "where cep = " & TabEND!CEP
                  Set TabCEP = DBARQEMP.OpenRecordset(SQL, 4)
                  If Not IsNull(TabCEP!Cidade) Then _
                     TabAUX!cidade_uf_emp = TabCEP!Cidade & " - " & TabCEP!UF
                  TabCEP.Close
               End If
            End If
         End If
         TabEND.Close

         'CLIENTE
         SQL = "select nome,cgccpf,razao_social from CLIENTE "
         SQL = SQL & "where cgccpf = '" & txtCNPJCPF.Text & "'"
         Set tabcli = DBARQEMP.OpenRecordset(SQL, 4)
         If Not tabcli.EOF Then
            CRITERIO = ""
            SQL = "select * from FONE "
            SQL = SQL & "where prop = '" & tabcli!CGCCPF & "'"
            Set TabEND = DBARQEMP.OpenRecordset(SQL, 4)
            While Not TabEND.EOF
               If Not IsNull(TabEND!DDD) Then _
                  CRITERIO = CRITERIO & "   (" & TabEND!DDD & ") "
               If Not IsNull(TabEND!Numero) Then _
                  CRITERIO = CRITERIO & TabEND!Numero
               TabEND.MoveNext
            Wend
            TabEND.Close

            TabAUX!nome_cli = tabcli!NOME
            If Not IsNull(tabcli!RAZAO_SOCIAL) Then _
               If tabcli!RAZAO_SOCIAL <> "" Then _
                  TabAUX!nome_cli = tabcli!RAZAO_SOCIAL

            'ENDEREÇO CLIENTE
            SQL = "select * from ENDERECO "
            SQL = SQL & " where prop = '" & txtCNPJCPF.Text & "'"
            If Len(tabcli!CGCCPF) <= 11 Then
               SQL = SQL & " and tipo = 'R' "
               Else: SQL = SQL & " and tipo = 'C' "
            End If
             Set TabEND = DBARQEMP.OpenRecordset(SQL, 4)
             If Not TabEND.EOF Then
                If Not IsNull(TabEND!Rua) Then _
                   TabAUX!end_cli = TabEND!Rua
                If Not IsNull(TabEND!Complemento) Then _
                   TabAUX!end_cli = TabAUX!end_cli & " ; " & TabEND!Complemento
                If Not IsNull(TabEND!Bairro) Then _
                   TabAUX!end_cli = TabAUX!end_cli & " ; " & TabEND!Bairro
                If Not IsNull(TabEND!CEP) Then
                   If TabEND!CEP <> "" Then
                      TabAUX!end_cli = TabAUX!end_cli & " ; " & TabEND!CEP
                      SQL = "select * from CEP "
                      SQL = SQL & "where cep = " & TabEND!CEP
                      Set TabCEP = DBARQEMP.OpenRecordset(SQL, 4)
                      If Not TabCEP.EOF Then
                         If Not IsNull(TabCEP!Cidade) Then _
                            TabAUX!end_cli = TabAUX!end_cli & " ; " & TabCEP!Cidade & " - " & TabCEP!UF
                      End If
                      TabCEP.Close
                   End If
                End If
             End If
             TabEND.Close
         End If
         tabcli.Close

         'VEICULO
         SQL = "select * from VEICULO "
         SQL = SQL & "where cgccpf = '" & txtCNPJCPF.Text & "'"
         SQL = SQL & " and placa = '" & Replace(txtPlaca.Text, "-", "") & "'"
         Set TabConsulta = DBARQAUX.OpenRecordset(SQL, 4)
         If Not TabConsulta.EOF Then
            TabAUX!descricao_veiculo = TabConsulta!Descricao
            TabAUX!ano_modelo = TabConsulta!Ano & "/" & TabConsulta!Modelo
            TabAUX!motor = Left(TabConsulta!motor, 30)
            TabAUX!chassi = TabConsulta!chassi
            TabAUX!placa = TabConsulta!placa
            TabAUX!KM = txtKM.Text
            'TABAUX!km = TABCONSULTA!km_atual
            'COR
            SQL = "select * from DESCR "
            SQL = SQL & "where tipo_a = 'Q' "
            SQL = SQL & "and codigo = " & TabConsulta!cor
            Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TabDESCR.EOF Then _
               TabAUX!cor = Trim(TabDESCR!desc_a)
            TabDESCR.Close
            'COMBUSTIVEL
            SQL = "select * from DESCR "
            SQL = SQL & "where tipo_a = 'S' "
            SQL = SQL & "and codigo = " & TabConsulta!combustivel
            Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TabDESCR.EOF Then _
               TabAUX!combustivel = Trim(TabDESCR!desc_a)
            TabDESCR.Close
         End If
         TabConsulta.Close
         TabAUX!FONE_cli = Trim(Left(CRITERIO, 50))
         TabAUX.Update
         TabAUX.Close
         
         'item serviço
         SQL = "select * from ITEMOS "
         SQL = SQL & "where numr_os = " & NUMR_OS
         Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
         While Not TabAUX.EOF
            SQL = "select * from RELOSITEM "
            SQL = SQL & "where numr_os = " & NUMR_OS
            SQL = SQL & " and tipo_item = 'S'" 'serviço
            SQL = SQL & " and codg_item = '" & TabAUX!Codg_tarefa & "'"
            Set TabConsulta = DBARQAUX.OpenRecordset(SQL)
            If Not TabConsulta.EOF Then
               TabConsulta.Edit
               Else: TabConsulta.AddNew
            End If

            TabConsulta!NUMR_OS = NUMR_OS
            TabConsulta!codg_item = TabAUX!Codg_tarefa

            SQL = "select descricao from TAREFA "
            SQL = SQL & "where codg_tarefa = '" & TabAUX!Codg_tarefa & "'"
            Set TabDESCR = DBARQAUX.OpenRecordset(SQL, 4)
            If Not TabDESCR.EOF Then _
               If Not IsNull(TabDESCR!Descricao) Then _
                  TabConsulta!desc_item = Left(TabDESCR!Descricao, 50)
            TabDESCR.Close
            
            SQL = "select nome from USUARIO "
            SQL = SQL & " where codigo = " & TabAUX!codg_mecanico
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TabUSU = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TabUSU.EOF Then _
               If Not IsNull(TabUSU!NOME) Then _
                  TabConsulta!mecanico = TabUSU!NOME
            TabUSU.Close

            TabConsulta!Valor_Item = TabAUX!valor_tarefa
            TabConsulta!desconto_item = TabAUX!valor_desc_tarefa
            TabConsulta!QTD = 1
            TabConsulta!tipo_item = "S"

            TabConsulta.Update
            TabConsulta.Close
            TabAUX.MoveNext
         Wend
         TabAUX.Close

            SQL = "select * from CABECAreq "
            SQL = SQL & " where numr_req = " & NUMR_OS
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TabCABECA = DBARQEMP.OpenRecordset(SQL)
            If Not TabCABECA.EOF Then
               TOTAL_PEÇAS_N = 0
               SQL = "select sum(valor_item*qtd_pedida) from ITEMREQ "
               SQL = SQL & " where numr_req = " & NUMR_OS
               SQL = SQL & " and empresa_id = " & EMPRESA_ID
               Set TabConsulta = DBARQEMP.OpenRecordset(SQL, 4)
               If Not TabConsulta.EOF Then _
                  If Not IsNull(TabConsulta.Fields(0).Value) Then _
                     TOTAL_PEÇAS_N = TabConsulta.Fields(0).Value
               TabConsulta.Close

               TabCABECA.Edit
                  TabCABECA!Valor_Desconto = VALOR_TOTAL_DESCONTO_N
                  TabCABECA!Valor_Total = TOTAL_PEÇAS_N
               TabCABECA.Update
            End If
            TabCABECA.Close
         
         'item peça
         SQL = "select * from ITEMREQ "
         SQL = SQL & " where numr_req = " & NUMR_OS
         SQL = SQL & " and empresa_id = " & EMPRESA_ID
         Set TabAUX = DBARQEMP.OpenRecordset(SQL, 4)
         While Not TabAUX.EOF
            SQL = "select * from RELOSITEM "
            SQL = SQL & "where numr_os = " & NUMR_OS
            SQL = SQL & " and tipo_item = 'P'" 'peças
            SQL = SQL & " and codg_item = '" & TabAUX!Codg_Prod & "'"
            Set TabConsulta = DBARQAUX.OpenRecordset(SQL)
            If Not TabConsulta.EOF Then
               TabConsulta.Edit
               Else: TabConsulta.AddNew
            End If

            TabConsulta!NUMR_OS = NUMR_OS
            TabConsulta!codg_item = TabAUX!Codg_Prod

            SQL = "select descricao from PRODUTO "
            SQL = SQL & " where codg_prod = '" & TabAUX!Codg_Prod & "'"
            SQL = SQL & " and empresa_id = " & EMPRESA_ID
            Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TabDESCR.EOF Then _
               If Not IsNull(TabDESCR!Descricao) Then _
                  TabConsulta!desc_item = Left(TabDESCR!Descricao, 50)
            TabDESCR.Close

            TabConsulta!Valor_Item = TabAUX!Valor_Item
            TabConsulta!desconto_item = (TabAUX!Valor_Item * TabAUX!QTD_PEDIDA) * TabAUX!PERC_desc / 100
            TabConsulta!QTD = TabAUX!QTD_PEDIDA
            TabConsulta!tipo_item = "P"

            TabConsulta.Update
            TabConsulta.Close
            TabAUX.MoveNext
         Wend
         TabAUX.Close
      End If
      TabTemp.Close
      DBARQAUX.Close

      frmINICIO.Dialogo.CancelError = True
      On Error GoTo TRATAERRO4
      'mostra a janela para impressora
      frmINICIO.Dialogo.ShowPrinter

      'IMPRIMIR RELATÓRIO
      frmINICIO.RELOS.SelectionFormula = ""
      frmINICIO.RELOS.Destination = 0
      frmINICIO.RELOS.SelectionFormula = "{RELOS.numr_os} = " & txtOs.Text
      frmINICIO.RELOS.ReportFileName = PATH_REL & "rel_abre_os.rpt"
      frmINICIO.RELOS.Action = 1

      SQL = "delete * from RELOS "
      SQL = SQL & "where numr_os = " & txtOs.Text
      DBARQAUX.Execute SQL

TRATAERRO4:
      Else: MsgBox "Empresa não cadastrada."
   End If
   TabEMP.Close
End Sub

Private Sub CABEÇALHO_IMPRESSÃO()
   Printer.Font = frmINICIO.Dialogo.FontSize
   Print #1, "----------------------------------------------------------------------------------------------"
   Print #1, TabEMP!Nome_Fant
   frmOSABRE.txtCGC.PromptInclude = False
      frmOSABRE.txtCGC.Text = TabEMP!CGC
   frmOSABRE.txtCGC.PromptInclude = True
   CRITERIO = "Insc.Estadual: " & Trim(TabEMP!IE)

   Print #1, Trim(TabEMP!RAZAO_SOCIAL); " - "; "CNPJ: "; Trim(frmOSABRE.txtCGC.Text); " - "; Trim(CRITERIO)
   Print #1, "----------------------------------------------------------------------------------------------"
   'procura endereço empresa
   SQL = "select * from ENDERECO e, CEP p "
   SQL = SQL & "where e.prop = '" & TabEMP!CGC & "'"
   SQL = SQL & " and e.cep = p.cep "
   SQL = SQL & " e.tipo = 'C' "
   Set TabEND = DBARQEMP.OpenRecordset(SQL, 4)
   If Not TabEND.EOF Then
      Print #1, Trim(TabEND!Rua); ", "; Trim(TabEND!Complemento); ", "; Trim(TabEND!Bairro); ", "; _
      Trim(TabEND.Fields("e.cep")); " "; Trim(TabEND!Cidade); " - "; Trim(TabEND!UF)
   End If
   TabEND.Close

   CRITERIO = ""
   SQL = "select * from FONE "
   SQL = SQL & "where prop = '" & TabEMP!CGC & "'"
   Set TabEND = DBARQEMP.OpenRecordset(SQL, 4)
   While Not TabEND.EOF
      If Not IsNull(TabEND!Numero) Then _
         CRITERIO = "(" & TabEND!DDD & ") "
      If Not IsNull(TabEND!Numero) Then _
         CRITERIO = CRITERIO & TabEND!Numero
      TabEND.MoveNext
   Wend
   TabEND.Close
   CRITERIO = "Telefax: " & CRITERIO
   Print #1,
      Printer.Font = frmINICIO.Dialogo.FontSize + 2
      Print #1, Spc(20); CRITERIO
   Print #1, "----------------------------------------------------------------------------------------------"
   Print #1,
   Print #1, Spc(10); "      LANTERNAGEM - PINTURA - ELETRICIDADE - INJEÇÃO"
   Print #1, Spc(10); "MECÂNICA EM GERAL - SUSPENSÃO - ALINHAMENTO - BALANCEAMENTO"
   Close #1
   LoadEXE ("C:\Arquivos de programas\Acessórios\WORDPAD.EXE c:\texte.txt")
End Sub


