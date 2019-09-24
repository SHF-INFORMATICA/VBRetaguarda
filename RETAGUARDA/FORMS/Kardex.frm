VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKardex 
   Caption         =   "Consulta Movimento Itens"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11040
   Icon            =   "Kardex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtNome 
         DataField       =   "Nome"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4260
         MaxLength       =   80
         TabIndex        =   12
         Top             =   315
         Width           =   6705
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3840
         Picture         =   "Kardex.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   810
         Width           =   285
      End
      Begin VB.TextBox txtDescricao 
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
         Height          =   360
         Left            =   4260
         MaxLength       =   29
         TabIndex        =   8
         Top             =   765
         Width           =   6705
      End
      Begin MSMask.MaskEdBox txtDtFim 
         Height          =   315
         Left            =   5310
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8880
         Top             =   -720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":6614
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":6A68
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":6D84
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":71D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":762C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":794C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":7DA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":80C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDtIni 
         Height          =   315
         Left            =   2220
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Height          =   375
         Left            =   2220
         TabIndex        =   13
         Top             =   315
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtProduto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   10
         ToolTipText     =   "Informe o código do produto."
         Top             =   765
         Width           =   1905
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
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
         Left            =   960
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
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
         Left            =   4260
         TabIndex        =   4
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor/Cliente:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1270
      ButtonWidth     =   3069
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rel. Consulta"
            Key             =   "print2"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   5
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
         Left            =   7440
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
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
               Picture         =   "Kardex.frx":8518
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":96B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":A741
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":B6F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Kardex.frx":C916
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin UltraGrid.SSUltraGrid GridEntrada 
      Height          =   2445
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4313
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   72679444
      RowConnectorColor=   -2147483633
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ValueLists      =   "Kardex.frx":DA21
      Bands           =   "Kardex.frx":DAA7
      Override        =   "Kardex.frx":F0BA
      Appearance      =   "Kardex.frx":F1A4
      CaptionAppearance=   "Kardex.frx":F1E0
      Caption         =   "Entrada"
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   11040
      DesignHeight    =   7455
   End
   Begin UltraGrid.SSUltraGrid GridSaida 
      Height          =   2445
      Left            =   0
      TabIndex        =   14
      Top             =   5040
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4313
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   72679444
      RowConnectorColor=   -2147483633
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ValueLists      =   "Kardex.frx":F21C
      Bands           =   "Kardex.frx":F2A2
      Override        =   "Kardex.frx":108B5
      Appearance      =   "Kardex.frx":1099F
      CaptionAppearance=   "Kardex.frx":109DB
      Caption         =   "Saída"
   End
End
Attribute VB_Name = "frmKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim rsGridEntrada As New ADODB.Recordset
   Dim rsGridSaida As New ADODB.Recordset

   Dim intCliFor As Integer

Private Sub Form_Load()
On Error GoTo ERRO_TRATA

   Call CentralizaJanela(frmKardex)
   Me.Caption = Me.Caption & Me.Name

   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtIni.Text = Date - 30
   txtDtFim.Text = Date
   txtDtIni.PromptInclude = True
   txtDtFim.PromptInclude = True

   ConsultaDados

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_load"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "consultar"
         ConsultaDados
         Exit Sub
      Case "print2"
         FORMULA_REL = "{Empresa.empresa_id} = " & EMPRESA_ID_N

         If txtProduto.Text <> "" Then _
            FORMULA_REL = FORMULA_REL & " and {SP_KARDEX;1.CODG_PROD} = '" & txtProduto.Text & "'"

         txtCGCCPF.PromptInclude = False

         If IsNumeric(txtCGCCPF.Text) Then _
            FORMULA_REL = FORMULA_REL & " and {SP_KARDEX;1.fornecedor_id} = " & FORNEC_ID_N

         txtCGCCPF.PromptInclude = True

         FORMULA_REL = FORMULA_REL & " and {SP_KARDEX;1.DT_ENTRADA} in date  (" & Format(txtDtIni.Text, "yyyy,MM,dd") & ") to date (" & Format(txtDtFim.Text, "yyyy,MM,dd") & ")"

         FORMULA_REL = FORMULA_REL & " and {SP_KARDEX;1.ENTRADA} = '" & "SAIDA" & "'"

         FORMULA_REL = FORMULA_REL & " and {SP_KARDEX;1.ENTRADA} = '" & "ENTRADA" & "'"

         If chkImp.Value = 1 Then _
            ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

         Nome_Relatorio = "rel_Kardex.rpt"
         frmRELATORIO10.Show 1

         LS_Envia_Parametro "{?@Codigo}", IIf(IsNumeric(txtCGCCPF.Text) = True, txtCGCCPF.Text, 0)
         LS_Envia_Parametro "{?@Dt_Inicial}", mdaI(txtDtIni.Text)
         LS_Envia_Parametro "{?@Dt_Final}", mdaF(txtDtFim.Text)
         'LS_Envia_Parametro "{?@tipo}", IIf(optTodas.Value = True, "", IIf(optVendas.Value = True, "SAIDA", "ENTRADA"))
         LS_Envia_Parametro "{?@Produto}", IIf(IsNumeric(txtProduto.Text) = True, txtProduto.Text, "")
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TELA
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdPesquisar_Click()
   CONSULTA_PRODUTO
End Sub

Private Sub txtProduto_LostFocus()
   If Trim(txtProduto.Text) <> "" Then
      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from PRODUTO "
      SQL = SQL & " where CODG_PRODUTO = '" & Trim(txtProduto.Text) & "'"
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabProduto.EOF Then
         MsgBox "Produto não Cadastrada.", vbOKOnly, "Atenção."
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)
         txtProduto.SetFocus
         Exit Sub
         Else: txtDescricao.Text = TabProduto!Descricao
      End If
   End If
End Sub

Private Sub txtproduto_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF7
         CONSULTA_PRODUTO
   End Select
End Sub

Private Sub TXTCGCCPF_GotFocus()
On Error GoTo ERRO_TRATA

   MOSTRA_RODAPE "ESC - SAIR", "F7 - Consulta Fornecedores", "", "", ""

   txtCGCCPF.PromptInclude = False
   If txtCGCCPF.Text = "" Then _
      txtCGCCPF.Mask = "##############"

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_GotFocus"
End Sub

Private Sub TXTCGCCPF_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         frmDISPLAYFORNECEDOR.Show 1
         If CNPJCPF_A <> "" Then _
            txtCGCCPF.Text = CNPJCPF_A
         CNPJCPF_A = ""
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCGCCPF_KeyDown"
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA
    Dim RstTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim dblTemp As Double

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtCGCCPF.PromptInclude = False
      If txtCGCCPF.Text = "" Then
         MsgBox "Informe CNPJ/CPF corretamente"
         txtCGCCPF.SetFocus
         Exit Sub
         Else
            If Len(txtCGCCPF.Text) > 0 Then
               Select Case Len(txtCGCCPF.Text)
                  Case Is = 11
                    If Not CALCULACPF(txtCGCCPF.Text) Then
                       MsgBox "CPF com DV incorreto !!!"
                       txtCGCCPF.PromptInclude = False
                       txtCGCCPF = ""
                       txtCGCCPF.SetFocus
                       Exit Sub
                    End If
                  Case Is = 14
                    If Not VALIDACGC(txtCGCCPF.Text) Then
                       MsgBox "CNPJ com DV incorreto !!! "
                       txtCGCCPF.PromptInclude = False
                       txtCGCCPF = ""
                       txtCGCCPF.SetFocus
                       Exit Sub
                    End If
                  Case Is > 14
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCGCCPF = ""
                     txtCGCCPF.SetFocus
                     Exit Sub
                  Case Is < 11
                     MsgBox "CNPJ/CPF com DV incorreto !!! "
                     txtCGCCPF = ""
                     txtCGCCPF.SetFocus
                     Exit Sub
               End Select
               Else
                  MsgBox "CNPJ/CPF com DV incorreto !!! "
                  txtCGCCPF = ""
                  txtCGCCPF.SetFocus
                  Exit Sub
            End If
            txtCGCCPF.PromptInclude = False
            CRITERIO = txtCGCCPF.Text
      End If
      txtCGCCPF.PromptInclude = False
      If txtCGCCPF.Text <> "" Then
         CRITERIO = txtCGCCPF.Text
         'txtCGCCPF.Mask = "##############"
         If Not IsNull(txtCGCCPF.Text) Then
            If Len(txtCGCCPF.Text) <= 11 Then _
               txtCGCCPF.Mask = "###.###.###-##"
            If Len(txtCGCCPF.Text) > 11 Then _
               txtCGCCPF.Mask = "##.###.###/####-##"
         End If
         txtCGCCPF.Text = CRITERIO
      End If
      txtCGCCPF.PromptInclude = False

      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where CGCCPF = '" & txtCGCCPF.Text & "'"
      If TabCliente.State = 1 Then TabCliente.Close
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCliente.EOF Then
         Beep
         MsgBox "Fornecedor não Cadastrado.", vbOKOnly, "Atenção !!!"
         txtCGCCPF.SetFocus
         Exit Sub
         Else
            If TabCliente!NOME <> "" Then _
               txtNome.Text = TabCliente!NOME
               FORNEC_ID_N = TabCliente!FORNECEDOR_ID
            If Not IsNull(TabCliente!Status) Then
               If TabCliente!Status <> "A" Then
                  MsgBox "Fornecedor Desativado, Favor Atualizar Cadastro!"
                  txtCGCCPF.SetFocus
                  Exit Sub
               End If
            End If
            If RstTemp.State = 1 Then RstTemp.Close
            RstTemp.Open "Select * From ENDERECO Where PROP='" & txtCGCCPF.Text & "'", CONECTA_RETAGUARDA, , , adCmdText
            
            If Not RstTemp.EOF Then
                'Pegou o CEP do cliente
                If Not IsNull(RstTemp!CEP) Then
                   dblTemp = RstTemp!CEP
                Else 'Não tem cadastrado cep, impossivel fazer tributacao sem a uf
                    RstTemp.Close
                    MsgBox "O Cadastro do Fornecedor não está completo. Verique os dados (CEP, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
                    txtCGCCPF.SetFocus
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
                       txtCGCCPF.SetFocus
                       Exit Sub
                    End If
                Else
                    RstTemp.Close
                    MsgBox "O Sistema verificou que esta empresa nao esta com os dados cadastrais incompletos. Verique-os, principalmente o Estado(UF) da empresa"
                    txtCGCCPF.SetFocus
                    Exit Sub
                End If
            Else
               RstTemp.Close
               MsgBox "O Sistema verificou que este Fornecedor esta com cadastrais incompletos."
               txtCGCCPF.SetFocus
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

Private Sub txtDtFim_LostFocus()
On Error GoTo ERRO_TRATA

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
On Error GoTo ERRO_TRATA

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
On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtFim.SetFocus
   End If
   Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTINI_KeyPress"
End Sub

Private Sub txtDtFim_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO_TRATA
   If KeyAscii = 13 Then
      KeyAscii = 0
      txtDtIni.SetFocus
   End If
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtDTfim_KeyPress"
End Sub

Private Sub LIMPA_TELA()
   txtNome.Text = ""
   txtCGCCPF.PromptInclude = False
   txtCGCCPF.Text = ""
   txtDtIni.PromptInclude = False
   txtDtFim.PromptInclude = False
   txtDtFim.Text = ""
   txtDtIni.Text = ""
   ConfiguraGrid GridEntrada
End Sub

Private Sub LE_NOTA()
   If TabNOTA.State = 1 Then TabNOTA.Close
   If IsNumeric(CRITERIO) Then
      TabNOTA.Open "Select * from NotaEntrada Where Empresa_id = " & EMPRESA_ID_N & " and Numr_pedido_compra = " & CRITERIO & "", CONECTA_RETAGUARDA, , , adCmdText
   Else
      TabNOTA.Open "Select * from NotaEntrada Where Empresa_id = " & EMPRESA_ID_N & "", CONECTA_RETAGUARDA, , , adCmdText
   End If
   If Not TabNOTA.EOF Then
      SQL = "select * from FORNECEDOR "
      SQL = SQL & " where fornecedor_id = " & TabNOTA!FORNECEDOR_ID
      If TabCliente.State = 1 Then TabCliente.Close
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         FORNEC_ID_N = TabNOTA!FORNECEDOR_ID
         txtCGCCPF.PromptInclude = False
         txtCGCCPF.Text = TabCliente!CGCCPF
         txtNome.Text = TabCliente!NOME
         txtCGCCPF.PromptInclude = True
      End If
      TabCliente.Close
      txtDtIni.PromptInclude = False
      txtDtFim.PromptInclude = False
      txtDtIni.Text = TabNOTA!DT_ENTRADA
      txtDtFim.Text = TabNOTA!DT_ENTRADA
      txtDtIni.PromptInclude = True
      txtDtFim.PromptInclude = True
   End If
   TabNOTA.Close
End Sub

Sub CONSULTA_PRODUTO()
   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""
End Sub

Sub ConsultaDados()
On Error GoTo ERRO_TRATA

   Msg = "Pesquisando..."

   GRID_ENTRADA

   GRID_SAIDA

Exit Sub
ERRO_TRATA:
   If Err.Number = 3705 Then 'objeto aberto
      rsGridEntrada.Close
   ElseIf Err.Number = 3704 Then 'objeto fechado
      'Resume Next
   End If
End Sub

Sub GRID_ENTRADA()
   ConfiguraGrid GridEntrada

   strRetorno = " DECLARE @Dt_Inicial Datetime"
   strRetorno = strRetorno & " DECLARE @Dt_Final Datetime"
   strRetorno = strRetorno & " DECLARE @Tipo varchar(50)"
   strRetorno = strRetorno & " DECLARE @Codigo int"
   strRetorno = strRetorno & " DECLARE @Produto varchar(50)"

   strRetorno = strRetorno & " SELECT @Dt_Inicial = '" & DMA(txtDtIni.Text) & "'"
   strRetorno = strRetorno & " SELECT @Dt_Final = '" & DMA(txtDtFim.Text) & "'"

   strRetorno = strRetorno & " SELECT @Tipo = 'ENTRADA'"

   If txtCGCCPF.Text <> "" Then
      strRetorno = strRetorno & " SELECT @Codigo = " & intCliFor
      Else: strRetorno = strRetorno & " SELECT @Codigo = " & 0
   End If

   If txtProduto.Text <> "" Then
      strRetorno = strRetorno & " SELECT @Produto = '" & txtProduto.Text & "'"
      Else: strRetorno = strRetorno & " SELECT @Produto = '" & "" & "'"
   End If

   If rsGridEntrada.State = 1 Then _
      rsGridEntrada.Close

   strRetorno = strRetorno & " EXEC sp_Kardex @Dt_Inicial, @Dt_Final, @tipo, @Codigo, @Produto "

   strConexaoGrid = "PROVIDER=MSDataShape;DATA PROVIDER=SQLOLEDB;SERVER=" & SERVIDOR_MEGASIM
   strConexaoGrid = strConexaoGrid & ";DATABASE=" & NOME_BANCO_DADOS
   strConexaoGrid = strConexaoGrid & ";UID=sa;"
   strConexaoGrid = strConexaoGrid & "PWD=" & SENHA_ADM_SQLSERVER
   strConexaoGrid = strConexaoGrid & ";CommandTimeout = 9000"

   rsGridEntrada.Open strRetorno, strConexaoGrid, , , adCmdText

   Set GridEntrada.DataSource = rsGridEntrada

   GridEntrada.Caption = "Entrada: " & rsGridEntrada.RecordCount
End Sub

Sub GRID_SAIDA()
   ConfiguraGrid GridSaida

   strRetorno = " DECLARE @Dt_Inicial Datetime"
   strRetorno = strRetorno & " DECLARE @Dt_Final Datetime"
   strRetorno = strRetorno & " DECLARE @Tipo varchar(50)"
   strRetorno = strRetorno & " DECLARE @Codigo int"
   strRetorno = strRetorno & " DECLARE @Produto varchar(50)"

   strRetorno = strRetorno & " SELECT @Dt_Inicial = '" & DMA(txtDtIni.Text) & "'"
   strRetorno = strRetorno & " SELECT @Dt_Final = '" & DMA(txtDtFim.Text) & "'"

   strRetorno = strRetorno & " SELECT @Tipo = 'SAIDA'"
   
   If txtCGCCPF.Text <> "" Then
      strRetorno = strRetorno & " SELECT @Codigo = " & intCliFor
      Else: strRetorno = strRetorno & " SELECT @Codigo = " & 0
   End If

   If txtProduto.Text <> "" Then
      strRetorno = strRetorno & " SELECT @Produto = '" & txtProduto.Text & "'"
      Else: strRetorno = strRetorno & " SELECT @Produto = '" & "" & "'"
   End If

   If rsGridSaida.State = 1 Then _
      rsGridSaida.Close

   strRetorno = strRetorno & " EXEC sp_Kardex @Dt_Inicial, @Dt_Final, @tipo, @Codigo, @Produto "

   strConexaoGrid = "PROVIDER=MSDataShape;DATA PROVIDER=SQLOLEDB;SERVER=" & SERVIDOR_MEGASIM
   strConexaoGrid = strConexaoGrid & ";DATABASE=" & NOME_BANCO_DADOS
   strConexaoGrid = strConexaoGrid & ";UID=sa;"
   strConexaoGrid = strConexaoGrid & "PWD=" & SENHA_ADM_SQLSERVER
   strConexaoGrid = strConexaoGrid & ";CommandTimeout = 9000"

   rsGridSaida.Open strRetorno, strConexaoGrid, , , adCmdText

   Set GridSaida.DataSource = rsGridSaida

   GridSaida.Caption = "saida: " & rsGridSaida.RecordCount
End Sub
