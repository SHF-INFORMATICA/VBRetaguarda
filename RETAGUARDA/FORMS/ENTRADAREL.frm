VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmENTRADAREL 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Entrada"
   ClientHeight    =   2835
   ClientLeft      =   3405
   ClientTop       =   2850
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ENTRADAREL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      Height          =   240
      Left            =   7560
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtNome 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Nome"
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      MaxLength       =   80
      TabIndex        =   4
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox txtNota 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1980
      MaxLength       =   9
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   3175
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
            Caption         =   "Rel. &Entrada"
            Key             =   "entrada"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rel Ent. &Itens"
            Key             =   "itens"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   5
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
               Picture         =   "ENTRADAREL.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ENTRADAREL.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ENTRADAREL.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ENTRADAREL.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ENTRADAREL.frx":A010
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   405
      Left            =   1980
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDtFim 
      Height          =   405
      Left            =   5430
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtCgcCpf 
      Height          =   405
      Left            =   1980
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4020
      TabIndex        =   8
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nº Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   6
      Top             =   1560
      Width           =   1425
   End
End
Attribute VB_Name = "frmENTRADAREL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Me.Caption = Me.Caption & Me.Name

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtIni.Text = Date
   txtDtFim.Text = Date
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_load"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         CRITERIO_A = ""
         frmNOTACONSULTA.Show 1
         If CRITERIO_A <> "" Then
            LE_NOTA
            txtDtIni.SetFocus
            Exit Sub
         End If
      Case "entrada"
         FORMULA_REL = "{NOTAENTRADA.estabelecimento_ID} = " & EMPRESA_ID_N
         If IsNumeric(txtNOTA.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.Numr_Nota} = " & txtNOTA.Text

         txtCNPJCPF.PromptInclude = False
         If IsNumeric(txtCNPJCPF.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.fornecedor_id} = " & FORNEC_ID_N
         txtCNPJCPF.PromptInclude = True

         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.DT_Entrada} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
         
         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Entrada.rpt"
         frmRELATORIO10.Show 1
      Case "itens"
         FORMULA_REL = "{NOTAENTRADA.estabelecimento_ID} = " & EMPRESA_ID_N
         If IsNumeric(txtNOTA.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.Numr_Nota} = " & txtNOTA.Text

         txtCNPJCPF.PromptInclude = False
         If IsNumeric(txtCNPJCPF.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.fornecedor_id} = " & FORNEC_ID_N
         txtCNPJCPF.PromptInclude = True

         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.DT_Entrada} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Entrada_Itens.rpt"
         frmRELATORIO10.Show 1
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TELA
         txtNOTA.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub TXTCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Fornecedores"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   
   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then
      txtCNPJCPF.Mask = "##############"
   End If

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
         If Trim(CNPJCPF_A) <> "" Then _
            txtCNPJCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyDown"
End Sub

Private Sub TXTCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

    Dim RstTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim dblTemp As Double

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text = "" Then
         MsgBox "Informe CNPJ/CPF corretamente"
         txtCNPJCPF.SetFocus
         Exit Sub
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
            txtCNPJCPF.PromptInclude = False
            CRITERIO_A = txtCNPJCPF.Text
      End If
      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text <> "" Then
         CRITERIO_A = txtCNPJCPF.Text
         'TXTCNPJCPF.Mask = "##############"
         If Not IsNull(txtCNPJCPF.Text) Then
            If Len(txtCNPJCPF.Text) <= 11 Then _
               txtCNPJCPF.Mask = "###.###.###-##"
            If Len(txtCNPJCPF.Text) > 11 Then _
               txtCNPJCPF.Mask = "##.###.###/####-##"
         End If
         txtCNPJCPF.Text = CRITERIO_A
      End If
      txtCNPJCPF.PromptInclude = False

      SQL = "select PESSOA.Descricao,fornecedor_id,status from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where cnpjcpf = '" & Trim(txtCNPJCPF.Text) & "'"
      If TabCliente.State = 1 Then TabCliente.Close
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "Fornecedor não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCNPJCPF.SetFocus
         Exit Sub
         Else
            If Trim(TabCliente.Fields("descricao").Value) <> "" Then _
               txtNome.Text = Trim(TabCliente.Fields("descricao").Value)
               FORNEC_ID_N = TabCliente!FORNECEDOR_ID
            If Not IsNull(TabCliente!STATUS) Then
               If TabCliente!STATUS <> "A" Then
                  MsgBox "Fornecedor Desativado, Favor Atualizar Cadastro!"
                  txtCNPJCPF.SetFocus
                  Exit Sub
               End If
            End If
            If RstTemp.State = 1 Then RstTemp.Close
            RstTemp.Open "select * from ENDERECO Where pessoa_id = " & PESSOA_ID_N, CONECTA_RETAGUARDA, , , adCmdText
            
            If Not RstTemp.EOF Then
                'Pegou o CEP do cliente
                If Not IsNull(RstTemp!CEP_ID) Then
                   dblTemp = RstTemp!CEP_ID
                Else 'Não tem cadastrado CEP_id, impossivel fazer tributacao sem a uf
                    RstTemp.Close
                    MsgBox "O Cadastro do Fornecedor não está completo. Verique os dados (CEP_id, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                    txtCNPJCPF.SetFocus
                    Exit Sub
                End If
                RstTemp.Close
                
                'Pegar a uf do cliente
                RstTemp.Open "select * from CEP where cep_ID=" & dblTemp, CONECTA_RETAGUARDA, , , adCmdText
                If Not RstTemp.EOF Then
                    If Not IsNull(RstTemp!UF) Then
                       RstTemp.Close
                    Else 'UF nao localizada
                       RstTemp.Close
                       MsgBox "O Cadastro do fornecedor não está completo. Verique os dados (CEP_id, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                       txtCNPJCPF.SetFocus
                       Exit Sub
                    End If
                Else
                    RstTemp.Close
                    MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais incompletos. Verique-os, principalmente o Estado(UF) da empresa"
                    txtCNPJCPF.SetFocus
                    Exit Sub
                End If
            Else
               RstTemp.Close
               MsgBox "O Sistema verificou que este Fornecedor esta com cadastrais incompletos."
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
      End If
      txtDtIni.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_KeyPress"
End Sub

Private Sub TXTDTFIM_GotFocus()
   txtDtFim.PromptInclude = True
End Sub

Private Sub txtDtFim_LostFocus()
'On Error GoTo ERRO_TRATA

   txtDtFim.PromptInclude = True
   If Not IsDate(txtDtFim.Text) Then
      txtDtFim.PromptInclude = False
         txtDtFim.Text = Date
      txtDtFim.PromptInclude = True
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDtfim_LostFocus"
End Sub

Private Sub TXTDTINI_GotFocus()
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

Private Sub txtDtIni_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      TabNOTA.Open "select Numr_Nota from NotaEntrada Where estabelecimento_id = " & ESTABELECIMENTO_ID_N & " and Numr_Nota = " & txtNOTA.Text & "", CONECTA_RETAGUARDA, , , adCmdText
      If TabNOTA.EOF Then
         MsgBox "Numero de Nota Inexistente, favor Conferir!", vbExclamation, ""
         txtNOTA.Text = ""
         txtNOTA.SetFocus
         TabNOTA.Close
         Exit Sub
         Else
            LE_NOTA
            Exit Sub
      End If
      If TabNOTA.State = 1 Then _
         TabNOTA.Close
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnota_KeyPress"
End Sub

Private Sub LIMPA_TELA()
'On Error GoTo ERRO_TRATA

   txtNOTA.Text = ""
   txtNome.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtNOTA.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TELA"
End Sub

Private Sub LE_NOTA()
'On Error GoTo ERRO_TRATA

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   If IsNumeric(CRITERIO_A) Then
      TabNOTA.Open "select * from NotaEntrada Where estabelecimento_id = " & ESTABELECIMENTO_ID_N & " and Numr_pedido_compra = " & CRITERIO_A & "", CONECTA_RETAGUARDA, , , adCmdText
      Else: TabNOTA.Open "select * from NotaEntrada Where estabelecimento_id = " & ESTABELECIMENTO_ID_N & " and Numr_Nota = " & txtNOTA.Text & "", CONECTA_RETAGUARDA, , , adCmdText
   End If
   If Not TabNOTA.EOF Then
      txtNOTA.Text = TabNOTA!NUMR_NOTA

      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close

      SQL = "select * from vwFornecedor WITH (NOLOCK)"
      SQL = SQL & " where fornecedor_id = " & TabNOTA!FORNECEDOR_ID
      TabFornecedor.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFornecedor.EOF Then
         FORNEC_ID_N = TabNOTA!FORNECEDOR_ID
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = "" & Trim(TabFornecedor.Fields("cnpjcpf").Value)
         txtNome.Text = "" & Trim(TabFornecedor.Fields("descricao").Value)
         txtCNPJCPF.PromptInclude = True
      End If
      If TabFornecedor.State = 1 Then _
         TabFornecedor.Close

      txtDtIni.PromptInclude = False
      txtDtFim.PromptInclude = False
      txtDtIni.Text = TabNOTA!DT_ENTRADA
      txtDtFim.Text = TabNOTA!DT_ENTRADA
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LE_NOTA"
End Sub
