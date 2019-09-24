VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPedidoClienteAcerto 
   Caption         =   "Acerto Pedido Cliente"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoClienteAcerto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   8325
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCliNovo 
      Enabled         =   0   'False
      Height          =   360
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   5415
   End
   Begin VB.TextBox txtCli 
      Enabled         =   0   'False
      Height          =   360
      Left            =   3060
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   360
      Left            =   300
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsCli 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   2280
      Picture         =   "PedidoClienteAcerto.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   405
   End
   Begin MSMask.MaskEdBox txtCGCCPF 
      Height          =   360
      Left            =   3060
      TabIndex        =   4
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   20
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
   Begin MSMask.MaskEdBox txtCliAlter 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
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
   Begin Threed.SSCommand cmdAtualiza 
      Height          =   615
      Left            =   6480
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      _Version        =   262144
      CaptionStyle    =   1
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Atualizar"
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      _Version        =   262144
      CaptionStyle    =   1
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Sair"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   8280
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Alterar para:"
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   8280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   8280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Acerto Pedido Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8370
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Pedido:"
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   240
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmPedidoClienteAcerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
   LIMPA_TUDO
   PEDIDO_ID_N = 0
   Unload Me
End Sub

Private Sub Form_Load()
   LIMPA_TUDO
   CARREGA_DADOS
End Sub

Private Sub txtCliAlter_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      TRATA_PESSOA txtCGCCPF.Text
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCliAlter_KeyPress"
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   txtCliAlter.PromptInclude = False
   txtCliAlter.Text = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1
   If Trim(CNPJCPF_A) <> "" Then
      txtCliAlter.Text = ""
      txtCliAlter.Mask = "##############"
      txtCliAlter.Text = CNPJCPF_A

      Call txtCliAlter_KeyPress(13)
   End If
   CNPJCPF_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsCli_Click"
End Sub

Private Sub cmdAtualiza_Click()
   GRAVA_ALTERAÇÃO
End Sub

Sub CARREGA_DADOS()
'On Error GoTo ERRO_TRATA

   If PEDIDO_ID_N > 0 Then
      txtPedido.Text = PEDIDO_ID_N

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select cliente_id,cgccpf,status,nome_cliente from PEDIDO "
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
      If Not TabConsulta.EOF Then
         txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = Trim(TabConsulta.Fields("cgccpf").Value)
         txtCNPJCPF.PromptInclude = False

         txtCli.Text = Trim(TabConsulta.Fields("nome_cliente").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_DADOS"
End Sub

Sub GRAVA_ALTERAÇÃO()
'On Error GoTo ERRO_TRATA

   If PEDIDO_ID_N > 0 And CLIENTE_ID_N > 0 Then
      Msg = "Confirma Alteração?"
      PERGUNTA Msg, vbYesNo + 32, "Atualização", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         SQL = "update PEDIDO set "
         SQL = SQL & " nome_cliente = '" & Trim(txtCliNovo.Text) & "'"
         SQL = SQL & ", cgccpf = '" & Trim(txtCliAlter.Text) & "'"
         SQL = SQL & ", cliente_id = " & CLIENTE_ID_N

         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         LIMPA_TUDO
         CARREGA_DADOS
         Unload Me
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_ALTERAÇÃO"
End Sub

Sub LIMPA_TUDO()
   PESSOA_ID_N = 0
   CLIENTE_ID_N = 0
   txtCliAlter.PromptInclude = False
   txtCliAlter.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtPedido.Text = ""
   txtCli.Text = ""
   txtCliNovo.Text = ""
End Sub
