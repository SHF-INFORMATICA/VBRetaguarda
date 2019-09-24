VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCADASTROAGENCIA 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Agencia Bancária"
   ClientHeight    =   2940
   ClientLeft      =   3180
   ClientTop       =   2790
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "cadastroagencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   9000
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastroagencia.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastroagencia.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastroagencia.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastroagencia.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastroagencia.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastroagencia.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cadastroagencia.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Salvar Informações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "matar"
            Object.ToolTipText     =   "Excluir Cadastro"
            ImageIndex      =   12
         EndProperty
      EndProperty
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
         Left            =   7560
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   9240
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   25
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":76BE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   7680
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":868E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":9A61
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":AAF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":BD58
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":CD0D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":E40A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":F832
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":1019A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":10A44
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":112EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":120C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadastroagencia.frx":13263
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   50
      TabIndex        =   14
      Top             =   720
      Width           =   8895
      Begin VB.TextBox txtBanco 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbBancoAux 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbBanco 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Banco:"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   50
      TabIndex        =   16
      Top             =   1440
      Width           =   8895
      Begin VB.TextBox txtComp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtRua 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtUF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtNomeAG 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   60
         TabIndex        =   3
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox txtAG 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtCep 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5040
         TabIndex        =   23
         Top             =   795
         Width           =   1260
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   1515
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rua/Avenida/Praça:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2040
         TabIndex        =   21
         Top             =   795
         Width           =   1710
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   20
         Top             =   1515
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   795
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8040
         TabIndex        =   18
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Agencia:"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   50
      TabIndex        =   24
      Top             =   2160
      Width           =   8895
      Begin VB.TextBox txtDDD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtL 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
      Begin MSComctlLib.Toolbar ToolbarFONE 
         Height          =   465
         Left            =   8280
         TabIndex        =   27
         Top             =   200
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   820
         ButtonWidth     =   847
         ButtonHeight    =   820
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "gravar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "matar"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DDD:"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Local:"
         Height          =   240
         Index           =   13
         Left            =   3360
         TabIndex        =   25
         Top             =   240
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmCADASTROAGENCIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmCADASTROAGENCIA)
   'LISTAFONE.CausesValidation = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Resize()
'On Error GoTo ERRO_TRATA

   NUMR_SEQ_N = 0

   If TabBANCO.State = 1 Then _
      TabBANCO.Close

   SQL = "select * from BANCO "
   SQL = SQL & " order by nome_banco"
   TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabBANCO.EOF
      cmbBanco.AddItem Trim(TabBANCO!Nome_Banco) & " - " & TabBANCO!Codg_Banco
      cmbBancoAux.AddItem TabBANCO.Fields("banco_id").Value
      TabBANCO.MoveNext
   Wend
   If TabBANCO.State = 1 Then _
      TabBANCO.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Resize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape: Unload Me
      Case vbKeyF9
         LIMPA_AG
         cmbBanco.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_KeyDown"
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
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_AG
         cmbBanco.SetFocus
      Case "gravar"
         If cmbBancoAux.Text <> "" And txtAG.Text <> "" Then
            GRAVA_AG
            LIMPA_BODY
            cmbBanco.SetFocus
            Exit Sub
         End If
      Case "print"
         FORMULA_REL = ""
         If IsNumeric(cmbBancoAux.Text) Then
            FORMULA_REL = "{AGENCIA.banco} = '" & cmbBancoAux.Text & "'"
            If txtAG.Text <> "" Then _
               FORMULA_REL = FORMULA_REL & " and " & "{agencia.numr_agencia} = " & "'" & txtAG.Text & "'"
         End If

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "relagen.rpt"
         frmRELATORIO10.Show 1
      Case "matar"
         If cmbBancoAux.Text <> "" And txtAG.Text <> "" Then
            Msg = "Confirma exclusão dessa Agencia ?"
            PERGUNTA Msg, vbYes, "Exclusao da Agencia", "DEMO.HLP", 1000
            If RESPOSTA = vbYes Then

               SQL = "delete  from AGENCIA "
               SQL = SQL & " where numr_agencia = '" & txtAG.Text & "'"
               CONECTA_RETAGUARDA.Execute SQL

               MsgBox "Opereção concluída com sucesso !!!"
               LIMPA_AG
               txtAG.SetFocus
            End If
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub ToolbarFONE_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "matar"
         If txtAG.Text <> "" And txtN.Text <> "" Then
            SQL = "delete from FONE "
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & "and NUMERO = '" & txtN.Text & "'"
            CONECTA_RETAGUARDA.Execute SQL
            CRITERIO_A = txtAG.Text

'            LISTAFONE.SetaGrid True, txtAG.Text, PATH_ARQ & NOME_BANCO_DADOS
            LIMPA_FONE
         End If
         txtDDD.SetFocus
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ToolbarFONE_ButtonClick"
End Sub

Private Sub txtBanco_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtBanco.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from Banco "
      SQL = SQL & " where codg_banco = '" & txtBanco.Text & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         cmbBanco.Text = Trim(TabTemp!Nome_Banco)
         cmbBancoAux.Text = Trim(TabTemp.Fields("banco_id").Value)
         txtBanco.Text = Trim(TabTemp.Fields("codg_banco").Value)
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBanco_LostFocus"
End Sub

Private Sub cmbBanco_LostFocus()
   If Trim(cmbBanco.Text) <> "" Then _
      MOSTRA_BANCO
End Sub

Private Sub cmbBanco_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Selecione um banco", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_GotFocus"
End Sub

Private Sub cmbBanco_Click()
'On Error GoTo ERRO_TRATA

   cmbBancoAux.ListIndex = cmbBanco.ListIndex
   SET_DATA_AGENCIA

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbBanco_Click"
End Sub

Private Sub txtUF_Change()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o estado"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtUF_Change"
End Sub

Private Sub txtAG_GotFocus()
'On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - Sair", "Informe número da agência", "", "", ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAG_GotFocus"
End Sub

Private Sub txtAG_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If txtAG.Text = "" Then
         MsgBox "Informe Código da Agencia !!!"
         txtAG.SetFocus
         Exit Sub
      End If
      CRITERIO_A = txtAG.Text
      SP_PROCURA_AG
      If Not TabAGENCIA.EOF Then
         KeyAscii = 0
         txtNomeAG.Text = TabAGENCIA!nome_agencia
         CRITERIO_A = TabAGENCIA!NUMR_AGENCIA

         'MOSTRA_ENDERECO

         'LISTAFONE.SetaGrid True, txtAG.Text, PATH_ARQ & NOME_BANCO_DADOS
      End If
      If TabAGENCIA.State = 1 Then _
         TabAGENCIA.Close

      txtNomeAG.SetFocus
      'cmbBanco.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAG_KeyPress"
End Sub

Private Sub txtBairro_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o bairro"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtBairro_GotFocus"
End Sub

Private Sub txtCep_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "F4 - Cadastra CEP"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "Informe número do CEP cadastrado"
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents

   txtCep.PromptInclude = False
   If txtCep.Text <> "" Then txtCep.Mask = "#####-###"
   txtCep.PromptInclude = True
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_GotFocus"
End Sub

Private Sub txtCep_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF4
         CRITERIO_A = ""
         frmCADASTROCEP.Show 1
         If CRITERIO_A <> "" Then
            If Not IsNull(CRITERIO_A) Then
               If Len(CRITERIO_A) <= 8 Then
                  txtCep.Mask = "#####-###"
               End If
               txtCep.PromptInclude = False
               txtCep.Text = CRITERIO_A
            End If
         End If
         CRITERIO_A = ""
         'txtCep.SetFocus
   End Select
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_KeyUp"
End Sub

Private Sub txtCep_LostFocus()
'On Error GoTo ERRO_TRATA

   txtCep.PromptInclude = False
   If txtCep.Text <> "" Then

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CEP "
      SQL = SQL & " where cep_ID = '" & txtCep.Text & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         MsgBox "Cep não cadastrado, F4-Cadastra Cep !!!"
         txtCep.PromptInclude = False
         txtCep.Text = ""
         txtCep.SetFocus
         Exit Sub
         Else
            txtCidade.Text = TabTemp!Cidade
            txtUF.Text = TabTemp!UF
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If
   txtCep.PromptInclude = True
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_LostFocus"
End Sub

Private Sub txtcidade_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe a cidade"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCidade_GotFocus"
End Sub

Private Sub txtComp_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o complemento"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtComp_GotFocus"
End Sub

Private Sub txtDDD_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o DDD"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDDD_GotFocus"
End Sub

Private Sub txtDDD_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtN.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDDD_KeyPress"
End Sub

Private Sub txtL_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o local do telefone"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtL_GotFocus"
End Sub

Private Sub txtN_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe o número do telefone"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtN_GotFocus"
End Sub

Private Sub txtN_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtL.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtN_KeyPress"
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      LIMPA_FONE
      txtDDD.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtl_KeyPress"
End Sub

Private Sub txtNomeAG_GotFocus()
'On Error GoTo ERRO_TRATA

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe nome da agência"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNomeAG_GotFocus"
End Sub

Private Sub txtnomeag_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If txtNomeAG.Text = "" Then
         MsgBox "Informe Nome da Agencia !!!"
         txtNomeAG.SetFocus
         Exit Sub
      End If
      KeyAscii = 0
      txtDDD.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtnomeag_KeyPress"
End Sub

Private Sub txtcep_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      txtCep.PromptInclude = False
      If txtCep.Text = "" Then
         txtCep.PromptInclude = True
         MsgBox "Informe Cep corretamente !!!"
         txtCep.SetFocus
         Exit Sub
      End If
      txtRua.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCep_KeyPress"
End Sub


Private Sub txtRua_GotFocus()
'On Error GoTo ERRO_TRATA
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents
   
   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = "Informe a rua/avenida/praça"
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtRua_GotFocus"
End Sub

Private Sub txtrua_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      If txtRua.Text <> "" Then
         KeyAscii = 0
         txtComp.SetFocus
      End If
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtrua_KeyPress"
End Sub

Private Sub txtcomp_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtBairro.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtcomp_KeyPress"
End Sub

Private Sub cmbbanco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtAG.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbbanco_KeyPress"
End Sub

Private Sub txtbanco_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      SendKeys "{tab}"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbbanco_KeyPress"
End Sub

Private Sub txtbairro_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      GRAVA_AG
      LIMPA_BODY
      SET_DATA_AGENCIA
      txtAG.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtbairro_KeyPress"
End Sub
'============================================================
Private Sub SET_DATA_AGENCIA()
'On Error GoTo ERRO_TRATA
   SQL = "select AGENCIA.*, CEP.*, ENDERECO.*"
   SQL = SQL & " from AGENCIA INNER JOIN "
   SQL = SQL & " (ENDERECO INNER JOIN CEP ON ENDERECO.CEP_ID = CEP.Cep_ID) "
   SQL = SQL & " ON AGENCIA.NUMR_AGENCIA = ENDERECO.PROP "
   SQL = SQL & " where codg_banco = '" & cmbBancoAux.Text & "'"
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SET_DATA_AGENCIA"
End Sub

Private Sub GRAVA_AG()
'On Error GoTo ERRO_TRATA

   If TabAGENCIA.State = 1 Then _
      TabAGENCIA.Close

   SQL = "select * from AGENCIA "
   SQL = SQL & " where numr_agencia = '" & Trim(txtAG.Text) & "'"
   SQL = SQL & " and banco_id = " & Trim(cmbBancoAux.Text)
   TabAGENCIA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabAGENCIA.EOF Then
      SQL = "UPDATE AGENCIA SET "
      SQL = SQL & " nome_agencia = '" & txtNomeAG.Text & "'"
      SQL = SQL & ", Banco_id = '" & cmbBancoAux.Text & "'"
      SQL = SQL & " where numr_agencia = '" & txtAG.Text & "'"
      Else
         SQL = "INSERT INTO AGENCIA "
            SQL = SQL & " (agencia_id,nome_agencia, NUMR_AGENCIA, Banco_id,codg_banco)"
         SQL = SQL & " VALUES ("
         SQL = SQL & MAX_ID("agencia_id", "agencia", "", "", "", "")
         SQL = SQL & ",'" & Trim(txtNomeAG.Text) & "'"
         SQL = SQL & ",'" & Trim(txtAG.Text) & "'"
         SQL = SQL & ",'" & Trim(cmbBancoAux.Text) & "'"
         SQL = SQL & ",'" & Trim(txtBanco.Text) & "'"
         SQL = SQL & " )"
   End If
   If TabAGENCIA.State = 1 Then _
      TabAGENCIA.Close

   CONECTA_RETAGUARDA.Execute SQL

   'GRAVA_ENDERECO
   SP_GRAVA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_AG"
End Sub

Private Sub LIMPA_AG()
'On Error GoTo ERRO_TRATA

   cmbBancoAux.Text = ""
   cmbBanco.Text = ""
   LIMPA_BODY
   LIMPA_FONE

   SQL = "select * from AGENCIA "
   SQL = SQL & " where numr_agencia=''"

'   LISTAFONE.SetaGrid True, txtAG.Text, PATH_ARQ & NOME_BANCO_DADOS

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_AG"
End Sub

Private Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtAG.Text = ""
   txtNomeAG.Text = ""
   txtCep.PromptInclude = False
   txtCep.Text = ""
   txtRua.Text = ""
   txtComp.Text = ""
   txtBairro.Text = ""
   txtCidade.Text = ""
   txtUF.Text = ""
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Private Sub LIMPA_FONE()
'On Error GoTo ERRO_TRATA

   txtN.Text = ""
   txtDDD.Text = ""
   txtL.Text = ""
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_FONE"
End Sub

Private Sub SP_PROCURA_AG()
'On Error GoTo ERRO_TRATA

   SQL = "select * from AGENCIA "
   SQL = SQL & " where numr_agencia = '" & Trim(txtAG.Text) & "'"
   SQL = SQL & " and banco_id = " & Trim(cmbBancoAux.Text)
   TabAGENCIA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SP_PROCURA_AG"
End Sub

Private Sub SP_PROCURA_FONETEMP()
'On Error GoTo ERRO_TRATA

   SP_GRAVA_FONE

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SP_PROCURA_FONETEMP"
End Sub

Private Sub SP_GRAVA_FONE()
'On Error GoTo ERRO_TRATA

   SQL = "select * from FONE "
   SQL = SQL & " where NUMERO = '" & txtN.Text & "'"
   SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
   TabFone.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabFone.EOF Then
      Do While Not TabFone.EOF
         SQL = "update FONE set "
         SQL = SQL & " NUMERO = '" & TabTemp!Numero & "'"
         SQL = SQL & ", ddd = " & TabTemp!ddd
         SQL = SQL & ", local = '" & TabTemp!local & "' "
         SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
         SQL = SQL & "and NUMERO = " & TabTemp!Numero
         CONECTA_RETAGUARDA.Execute SQL
         TabFone.MoveNext
      Loop
      Else
         If txtN.Text <> "" Then
            SQL = "insert into FONE values ("
            SQL = SQL & "'" & txtAG.Text & "'"
            SQL = SQL & ",'" & txtN.Text & "'"
            SQL = SQL & ",'" & txtDDD.Text & "'"
            SQL = SQL & ",'" & Trim(txtL.Text) & "' "
         End If
   End If
   If TabFone.State = 1 Then _
      TabFone.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SP_GRAVA_FONE"
End Sub
'==============================================
Private Sub MOSTRA_ENDERECO()
'On Error GoTo ERRO_TRATA

   CRITERIO_A = txtAG.Text
'implementar
   BUSCA_ENDERECO_PESSOA "C", ""  'endereço COMERCIAL
   If Not tabEndereco.EOF Then
      txtRua.Text = tabEndereco!Rua
      txtBairro.Text = tabEndereco!Bairro
      If Not IsNull(tabEndereco!Complemento) Then
         txtComp.Text = tabEndereco!Complemento
      End If
      If Not IsNull(tabEndereco!CEP_id) Then
         If tabEndereco!CEP_id <> "" Then
            txtCep.Text = tabEndereco!CEP_id

            If TabConsulta.State = 1 Then _
               TabConsulta.Close

            SQL = "select * from CEP "
            SQL = SQL & " where cep_ID = '" & tabEndereco!CEP_id & "'"
            TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabConsulta.EOF Then
               txtCep.Text = tabEndereco!CEP_id
               txtCidade.Text = TabConsulta!Cidade
               txtUF.Text = TabConsulta!UF
            End If
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
         End If
      End If
   End If
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ENDERECO"
End Sub

Private Sub GRAVA_ENDERECO()
'On Error GoTo ERRO_TRATA

   txtCep.PromptInclude = False
   sp_Grava_Endereco txtCep.Text, txtRua.Text, txtBairro.Text, txtComp.Text, "C", "0"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ENDERECO"
End Sub

Sub MOSTRA_BANCO()
'On Error GoTo ERRO_TRATA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If Trim(cmbBanco.Text) <> "" Then
      SQL = "select * from Banco "

      If IsNumeric(cmbBanco.Text) Then
         SQL = SQL & " where codg_banco = '" & cmbBanco.Text & "'"
         Else: SQL = SQL & " where nome_banco = '" & Trim(cmbBanco.Text) & "'"
      End If

      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         cmbBanco.Text = Trim(TabTemp!Nome_Banco)
         cmbBancoAux.Text = Trim(TabTemp.Fields("banco_id").Value)
         txtBanco.Text = Trim(TabTemp.Fields("codg_banco").Value)
      End If
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_BANCO"
End Sub
