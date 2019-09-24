VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelEntrada 
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNome 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataField       =   "Nome"
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      MaxLength       =   80
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   0
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
               Picture         =   "frmRelEntrada.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRelEntrada.frx":119A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRelEntrada.frx":2229
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRelEntrada.frx":31DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRelEntrada.frx":43FE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtDtIni 
      Height          =   405
      Left            =   1980
      TabIndex        =   3
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
      TabIndex        =   4
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
      TabIndex        =   5
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
Attribute VB_Name = "frmRelEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmRelEntrada)
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
         CRITERIO = ""
         frmNOTAENTRADACONSULTA.Show 1
         If CRITERIO <> "" Then
            LE_NOTA
            txtDtIni.SetFocus
            Exit Sub
         End If
      Case "entrada"
         FORMULA_REL = "{NOTAENTRADA.empresa_id} = " & EMPRESA_ID_N
         If IsNumeric(txtNota.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.Numr_Nota} = " & txtNota.Text

         txtCgcCpf.PromptInclude = False
         If IsNumeric(txtCgcCpf.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.fornecedor_id} = " & FORNEC_ID_N
         txtCgcCpf.PromptInclude = True

         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.DT_Entrada} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"
         
         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
         Nome_Relatorio = "rel_Entrada.rpt"
         frmRELATORIO10.Show 1
      Case "itens"
         FORMULA_REL = "{NOTAENTRADA.empresa_id} = " & EMPRESA_ID_N
         If IsNumeric(txtNota.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.Numr_Nota} = " & txtNota.Text

         txtCgcCpf.PromptInclude = False
         If IsNumeric(txtCgcCpf.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.fornecedor_id} = " & FORNEC_ID_N
         txtCgcCpf.PromptInclude = True

         If IsDate(txtDtIni.Text) And IsDate(txtDtFim.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {NOTAENTRADA.DT_Entrada} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"

         ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
         Nome_Relatorio = "rel_Entrada_Itens.rpt"
         frmRELATORIO10.Show 1
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TELA
         txtNota.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub TXTCGCCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - SAIR"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F7 - Consulta Fornecedores"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   
   txtCgcCpf.PromptInclude = False
   If txtCgcCpf.Text = "" Then
      txtCgcCpf.Mask = "##############"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_GotFocus"
End Sub

Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         If CPF_N <> "" Then _
            txtCgcCpf.Text = CPF_N
         CPF_N = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_KeyDown"
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
    Dim RstTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim dblTemp As Double

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCgcCpf.PromptInclude = False
      If txtCgcCpf.Text = "" Then
         MsgBox "Informe CNPJ/CPF corretamente"
         txtCgcCpf.SetFocus
         Exit Sub
         Else
            If Len(txtCgcCpf.Text) > 0 Then
               Select Case Len(txtCgcCpf.Text)
                  Case Is = 11
                    If Not CALCULACPF(txtCgcCpf.Text) Then
                       MsgBox "CPF com DV incorreto !!!"
                       txtCgcCpf.PromptInclude = False
                       txtCgcCpf = ""
                       txtCgcCpf.SetFocus
                       Exit Sub
                    End If
                  Case Is = 14
                    If Not VALIDACGC(txtCgcCpf.Text) Then
                       MsgBox "CNPJ com DV incorreto !!! "
                       txtCgcCpf.PromptInclude = False
                       txtCgcCpf = ""
                       txtCgcCpf.SetFocus
                       Exit Sub
                    End If
                  Case Is > 14
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCgcCpf = ""
                     txtCgcCpf.SetFocus
                     Exit Sub
                  Case Is < 11
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCgcCpf = ""
                     txtCgcCpf.SetFocus
                     Exit Sub
               End Select
               Else
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtCgcCpf = ""
                  txtCgcCpf.SetFocus
                  Exit Sub
            End If
            txtCgcCpf.PromptInclude = False
            CRITERIO = txtCgcCpf.Text
      End If
      txtCgcCpf.PromptInclude = False
      If txtCgcCpf.Text <> "" Then
         CRITERIO = txtCgcCpf.Text
         'txtCGCCPF.Mask = "##############"
         If Not IsNull(txtCgcCpf.Text) Then
            If Len(txtCgcCpf.Text) <= 11 Then _
               txtCgcCpf.Mask = "###.###.###-##"
            If Len(txtCgcCpf.Text) > 11 Then _
               txtCgcCpf.Mask = "##.###.###/####-##"
         End If
         txtCgcCpf.Text = CRITERIO
      End If
      txtCgcCpf.PromptInclude = False

      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where CGCCPF = '" & txtCgcCpf.Text & "'"
      If TabCliente.State = 1 Then TabCliente.Close
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "Fornecedor não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCgcCpf.SetFocus
         Exit Sub
         Else
            If TabCliente!NOME <> "" Then _
               txtNome.Text = TabCliente!NOME
               FORNEC_ID_N = TabCliente!FORNECEDOR_ID
            If Not IsNull(TabCliente!Status) Then
               If TabCliente!Status <> "A" Then
                  MsgBox "Fornecedor Desativado, Favor Atualizar Cadastro!"
                  txtCgcCpf.SetFocus
                  Exit Sub
               End If
            End If
            If RstTemp.State = 1 Then RstTemp.Close
            RstTemp.Open "Select * From ENDERECO Where PROP='" & txtCgcCpf.Text & "'", CONECTA_RETAGUARDA, , , adCmdText
            
            If Not RstTemp.EOF Then
                'Pegou o CEP do cliente
                If Not IsNull(RstTemp!CEP) Then
                   dblTemp = RstTemp!CEP
                Else 'Não tem cadastrado cep, impossivel fazer tributacao sem a uf
                    RstTemp.Close
                    MsgBox "O Cadastro do Fornecedor não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                    txtCgcCpf.SetFocus
                    Exit Sub
                End If
                RstTemp.Close
                
                'Pegar a uf do cliente
                RstTemp.Open "Select * From CEP Where CEP=" & dblTemp, CONECTA_RETAGUARDA, , , adCmdText
                If Not RstTemp.EOF Then
                    If Not IsNull(RstTemp!UF) Then
                       RstTemp.Close
                    Else 'UF nao localizada
                       RstTemp.Close
                       MsgBox "O Cadastro do fornecedor não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                       txtCgcCpf.SetFocus
                       Exit Sub
                    End If
                Else
                    RstTemp.Close
                    MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais incompletos. Verique-os, principalmente o Estado(UF) da empresa"
                    txtCgcCpf.SetFocus
                    Exit Sub
                End If
            Else
               RstTemp.Close
               MsgBox "O Sistema verificou que este Fornecedor esta com cadastrais incompletos."
               txtCgcCpf.SetFocus
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
   TRATA_ERROS Err.Description, Me.Name, "txtCGCCPF_KeyPress"
End Sub
Private Sub txtDTfim_GotFocus()
   txtDtFim.PromptInclude = True
End Sub

Private Sub txtDtfim_LostFocus()
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
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDTfim_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub txtnota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      TabNOTA.Open "Select Numr_Nota from NotaEntrada Where Empresa_id = " & EMPRESA_ID_N & " and Numr_Nota = " & txtNota.Text & "", CONECTA_RETAGUARDA, , , adCmdText
      If TabNOTA.EOF Then
         MsgBox "Numero de Nota Inexistente, favor Conferir!", vbExclamation, ""
         txtNota.Text = ""
         txtNota.SetFocus
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
   txtNota.Text = ""
   txtNome.Text = ""
   txtCgcCpf.PromptInclude = False
   txtCgcCpf.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   txtNota.SetFocus
End Sub

Private Sub LE_NOTA()
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   If IsNumeric(CRITERIO) Then
      TabNOTA.Open "Select * from NotaEntrada Where Empresa_id = " & EMPRESA_ID_N & " and Numr_pedido_compra = " & CRITERIO & "", CONECTA_RETAGUARDA, , , adCmdText
      Else: TabNOTA.Open "Select * from NotaEntrada Where Empresa_id = " & EMPRESA_ID_N & " and Numr_Nota = " & txtNota.Text & "", CONECTA_RETAGUARDA, , , adCmdText
   End If
   If Not TabNOTA.EOF Then
      txtNota.Text = TabNOTA!NUMR_NOTA

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where fornecedor_id = " & TabNOTA!FORNECEDOR_ID
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         FORNEC_ID_N = TabNOTA!FORNECEDOR_ID
         txtCgcCpf.PromptInclude = False
         txtCgcCpf.Text = TabCliente!CGCCPF
         txtNome.Text = TabCliente!NOME
         txtCgcCpf.PromptInclude = True
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      txtDtIni.PromptInclude = False
      txtDtFim.PromptInclude = False
      txtDtIni.Text = TabNOTA!DT_ENTRADA
      txtDtFim.Text = TabNOTA!DT_ENTRADA
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close
End Sub
