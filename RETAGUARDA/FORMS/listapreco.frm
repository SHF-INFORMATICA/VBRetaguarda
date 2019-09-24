VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLISTAPRECO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Preço"
   ClientHeight    =   2520
   ClientLeft      =   3405
   ClientTop       =   2565
   ClientWidth     =   5895
   Icon            =   "listapreco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5895
   Begin VB.ComboBox cmbFamiliaAux 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFORNEC 
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
      MaxLength       =   100
      TabIndex        =   1
      ToolTipText     =   "Fornecedor deste produto"
      Top             =   1440
      Width           =   3615
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
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtPerc 
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
      TabIndex        =   2
      Top             =   2040
      Width           =   615
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   3016
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
            Caption         =   "Rel. &Entrada"
            Key             =   "entrada"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
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
               Picture         =   "listapreco.frx":47C4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "listapreco.frx":48DE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "listapreco.frx":49E73
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "listapreco.frx":4AE28
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "listapreco.frx":4C048
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblfornec 
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
      Left            =   630
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label lblperc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Percentual:"
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
      Left            =   735
      TabIndex        =   3
      Top             =   2040
      Width           =   1320
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
         txtFORNEC.Text = ""
         txtperc.Text = ""
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

   cmbFamiliaAux.ListIndex = cmbFamilia.ListIndex

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbFamilia_Click"
End Sub

Private Sub txtfornec_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         CPF_N = ""
         frmDISPLAYFORNECEDOR.Show 1
         If CPF_N <> "" Then
            If TabFORNECEDOR.State = 1 Then _
               TabFORNECEDOR.Close

            SQL = "select nome,razao_social, fornecedor_id from FORNECEDOR "
            SQL = SQL & " where CGCCPF = '" & CPF_N & "'"
            TabFORNECEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabFORNECEDOR.EOF Then
               If Trim(TabFORNECEDOR!NOME) = "" Then
                  txtFORNEC.Text = Trim(TabFORNECEDOR!RAZAO_SOCIAL)
                  Else: txtFORNEC.Text = Trim(TabFORNECEDOR!NOME)
               End If
               FORNEC_ID_N = TabFORNECEDOR!FORNECEDOR_ID
            End If
            If TabFORNECEDOR.State = 1 Then _
               TabFORNECEDOR.Close
         End If
         CPF_N = ""
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

   If txtFORNEC.Text <> "" Then _
      FORMULA_REL = FORMULA_REL & " and {PRODUTO.fornecedor_id} = " & FORNEC_ID_N

   If Trim(txtperc.Text) = "" Then _
      txtperc.Text = 0
'crxReport.ParameterFields(0).ParameterType = "Percentual; " & txtperc.Text & ";True"

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

         cmbFamilia.AddItem Trim(TabTemp!Descricao) & "-" & TabTemp!familiaproduto_id
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
