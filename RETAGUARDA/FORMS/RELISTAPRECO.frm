VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLISTAPRECO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Preço"
   ClientHeight    =   2355
   ClientLeft      =   3405
   ClientTop       =   2565
   ClientWidth     =   5925
   Icon            =   "RELISTAPRECO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5925
   Begin VB.TextBox txtRazao 
      DataField       =   "Nome"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2535
      MaxLength       =   80
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CommandButton cmdConsulta 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   2085
      Picture         =   "RELISTAPRECO.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   405
   End
   Begin VB.ComboBox cmbFamiliaAux 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbFamilia 
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
      Left            =   2175
      TabIndex        =   0
      ToolTipText     =   "Selecione o grupo do produto."
      Top             =   960
      Width           =   3615
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   2
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
            Key             =   "entrada"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4080
         TabIndex        =   4
         Top             =   360
         Width           =   1455
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
               Picture         =   "RELISTAPRECO.frx":0F8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RELISTAPRECO.frx":2126
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RELISTAPRECO.frx":31B5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RELISTAPRECO.frx":416A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RELISTAPRECO.frx":538A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtCNPJCPF 
      Height          =   405
      Left            =   135
      TabIndex        =   7
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
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label lblgrupo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Família Produto:"
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
      Left            =   165
      TabIndex        =   1
      Top             =   960
      Width           =   1890
   End
End
Attribute VB_Name = "FrmLISTAPRECO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(FrmLISTAPRECO)
   Me.Caption = Me.Caption & " - " & Me.Name
   preencheComboGRUPO cmbFamilia

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
         cmbFamilia.Text = ""
         cmbFamiliaAux.Text = ""
         FORNEC_ID_N = 0
         txtCNPJCPF.PromptInclude = False
         txtCNPJCPF.Text = ""
         txtRazao.Text = ""
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

Private Sub cmdConsulta_Click()
'On Error GoTo ERRO_TRATA

   CNPJCPF_A = ""
   frmDISPLAYFORNECEDOR.Show 1
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

Private Sub cmbFamilia_Click()
'On Error GoTo ERRO_TRATA

   cmbFamiliaAux.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub txtfornec_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CNPJCPF_A = ""
         frmDISPLAYFORNECEDOR.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            If TabFORNECEDOR.State = 1 Then _
               TabFORNECEDOR.Close

            SQL = "select nome,razao_social, fornecedor_id from FORNECEDOR "
            SQL = SQL & " where CGCCPF = '" & CNPJCPF_A & "'"
            TabFORNECEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabFORNECEDOR.EOF Then
               If Trim(TabFORNECEDOR!NOME) = "" Then
                  txtFORNEC.Text = Trim(TabFORNECEDOR!razao_social)
                  Else: txtFORNEC.Text = Trim(TabFORNECEDOR!NOME)
               End If
               FORNEC_ID_N = TabFORNECEDOR!FORNECEDOR_ID
            End If
            If TabFORNECEDOR.State = 1 Then _
               TabFORNECEDOR.Close
         End If
         CNPJCPF_A = ""
         txtFORNEC.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtfornec_KeyDown"
End Sub

Sub IMPRIMIR_REL()
'On Error GoTo ERRO_TRATA

   FORMULA_REL = "{PRODUTO.TIPO_PROD} = '1'"
   FORMULA_REL = FORMULA_REL & " and {PRODUTO.situacao} = 'A'"

   If cmbFamilia.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.familiaproduto_id} = " & numeros(cmbFamiliaAux.Text)

   If FORNEC_ID_N > 0 Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.fornecedor_id} = " & FORNEC_ID_N

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS
   
   Nome_Relatorio = "rel_listapreco.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPRIMIR_REL"
End Sub

Private Sub preencheComboGRUPO(NomeCombo As ComboBox)
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

         cmbFamilia.AddItem Trim(TabTemp!Descricao) & "-" & TabTemp!codg_familia
         cmbFamiliaAux.AddItem TabTemp!familiaproduto_id
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


Private Sub txtCNPJCPF_GotFocus()
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

Private Sub txtcnpjcpf_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
            txtCNPJCPF.PromptInclude = True
            txtCNPJCPF.SetFocus
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub txtcnpjcpf_KeyPress(KeyAscii As Integer)
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
              CRITERIO = txtCNPJCPF.Text
        End If
        txtCNPJCPF.PromptInclude = False
        If Trim(txtCNPJCPF.Text) <> "" Then
           CRITERIO = txtCNPJCPF.Text
           If Not IsNull(txtCNPJCPF.Text) Then
              If Len(txtCNPJCPF.Text) <= 11 Then
                 txtCNPJCPF.Mask = "###.###.###-##"
                 Else: txtCNPJCPF.Mask = "##.###.###/####-##"
              End If
           End If
           txtCNPJCPF.Text = CRITERIO
           Else: txtCNPJCPF.Mask = "##############"
        End If
        txtCNPJCPF.PromptInclude = False

         If TabCliente.State = 1 Then _
            TabCliente.Close

         SQL = "select * from FORNECEDOR "
         SQL = SQL & " where CGCCPF = '" & txtCNPJCPF.Text & "'"
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
