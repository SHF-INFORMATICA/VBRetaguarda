VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOrcamento 
   Caption         =   "Orçamento Cliente"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
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
   Icon            =   "OrcCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "OrcCliente.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCliente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtValorDig"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "OrcCliente.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.TextBox txtValorDig 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9240
         TabIndex        =   34
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   60
         TabIndex        =   18
         Top             =   1440
         Width           =   10815
         Begin VB.TextBox txtUN 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   360
            Left            =   10320
            TabIndex        =   33
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtSeq 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   360
            Left            =   0
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtTotItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   9120
            MaxLength       =   12
            TabIndex        =   30
            ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtVlrUnitario 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   9360
            MaxLength       =   12
            TabIndex        =   28
            ToolTipText     =   "Informe o valor unitário do item ou aceite o valor informado pelo sistema."
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtLargura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   27
            ToolTipText     =   "Informe a quantidade de venda deste produto."
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtAltura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1080
            MaxLength       =   8
            TabIndex        =   26
            ToolTipText     =   "Informe a quantidade de venda deste produto."
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtQtde 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   22
            ToolTipText     =   "Informe a quantidade de venda deste produto."
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtProduto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1080
            TabIndex        =   1
            ToolTipText     =   "Informe o código do produto, F6-Excluir, F7-Consultar"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtDescricao 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3120
            MaxLength       =   29
            TabIndex        =   20
            Top             =   240
            Width           =   7095
         End
         Begin VB.CommandButton cmdConsProd 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2520
            Picture         =   "OrcCliente.frx":5C4A
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Vlr.Total = "
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7995
            TabIndex        =   31
            Top             =   1320
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Vlr.Unit.= "
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   8400
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblLargura 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Largura = "
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2865
            TabIndex        =   25
            Top             =   720
            Width           =   990
         End
         Begin VB.Label lblAltura 
            Alignment       =   1  'Right Justify
            Caption         =   "Altura = "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Qtde ="
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5880
            TabIndex        =   23
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Produto:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraCliente 
         Height          =   1215
         Left            =   60
         TabIndex        =   7
         Top             =   360
         Width           =   10815
         Begin VB.ComboBox cmbVendedorAUX 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3240
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cmdEstagioAUX 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5640
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cmbVendedor 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3240
            TabIndex        =   3
            ToolTipText     =   "Selecione um vendedor"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmdEstagio 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   5640
            TabIndex        =   4
            ToolTipText     =   "Selecione um vendedor"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPedido 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   2
            ToolTipText     =   "<Enter> Gera uma requisição nova ou informe o número de uma requisição já existente."
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtNome 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3480
            MaxLength       =   100
            TabIndex        =   10
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmdConsCli 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3000
            Picture         =   "OrcCliente.frx":664C
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   720
            Width           =   495
         End
         Begin MSMask.MaskEdBox txtCNPJCPF 
            Height          =   375
            Left            =   1080
            TabIndex        =   0
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
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
         Begin MSMask.MaskEdBox txtDtEmis 
            Height          =   360
            Left            =   9480
            TabIndex        =   39
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483637
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
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total = "
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   9720
            TabIndex        =   17
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblSituacao 
            AutoSize        =   -1  'True
            Caption         =   "Situação"
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   7320
            TabIndex        =   16
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2145
            TabIndex        =   13
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Estágio:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4785
            TabIndex        =   12
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblCli 
            Alignment       =   1  'Right Justify
            Caption         =   "Cliente:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblPedido 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3375
         Left            =   60
         TabIndex        =   35
         Top             =   2640
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5953
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
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   2
      MaxFontSize     =   12
      DesignWidth     =   10965
      DesignHeight    =   7350
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   720
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   1270
      ButtonWidth     =   3201
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
            Caption         =   "&Gravar"
            Key             =   "gravar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Consultar"
            Key             =   "consultar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cad. Cliente"
            Key             =   "CadCliente"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   9480
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10080
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":704E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":81E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":9277
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":A22C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":B337
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":C48D
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":C8DF
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":E756
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":FE0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OrcCliente.frx":11DEE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orçamento Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   11010
   End
End
Attribute VB_Name = "frmOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   INICIA_TELA

'   ABRE_PEDIDO
   'REMOVE_MENU

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         LIMPA_TUDO
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description & qual_tecla, Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If INDR_PEDIDO_VALIDO = True And Trim(txtPedido.Text) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select status from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & txtPedido.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields(0).Value) Then
            If TabTemp.Fields(0).Value <= 2 Then
               Msg = "Pedido pendente, deseja realmente cancelar essa venda ?"
               PERGUNTA Msg, vbYesNo + 32, "Cancelar", "DEMO.HLP", 1000
               If RESPOSTA = vbYes Then
                  If TabTemp.State = 1 Then _
                     TabTemp.Close
   
                  SQL = "delete from ITEMLANCAMENTO where numr_doc = " & txtPedido.Text
                  CONECTA_RETAGUARDA.Execute SQL
                  SQL = "delete from LANCAMENTO where numr_doc = " & txtPedido.Text
                  CONECTA_RETAGUARDA.Execute SQL
                  SQL = "delete from pedidotemp where pedido_id = " & txtPedido.Text
                  CONECTA_RETAGUARDA.Execute SQL

                  SQL = "update pedido set "
                  SQL = SQL & " status = 9 "
                  SQL = SQL & " where pedido_id = " & txtPedido.Text
                  CONECTA_RETAGUARDA.Execute SQL

                  INDR_PEDIDO_VALIDO = False
                  Else: Cancel = 1
               End If
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   End If

   If TRAZ_TIPO_USUARIO = 1 Then _
      End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "gravar"
         VAI_VENDA
      Case "consultar"
         CRITERIO_A = ""
         CNPJCPF_A = ""
         frmPedidoConsulta.Show 1
         If PEDIDO_ID_N > 0 Then
            Dim NUMR_PEDIDO_N As Long

            NUMR_PEDIDO_N = PEDIDO_ID_N

            LIMPA_TUDO
            txtPedido.Text = NUMR_PEDIDO_N
            CRITERIO_A = ""
            NUMR_PEDIDO_N = 0

            ABRE_PEDIDO
         End If
      Case "print"
         'GERA_IMPRESSAO
      Case "limpar"
         LIMPA_ORÇAMENTO
      Case "voltar"
         Unload Me
      Case "CadCliente"
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaCadastro.Show 1

          'frmCADASTROCLIENTE.Show 1
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdConsCli_Click()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   TIPO_PESSOA_CADASTRO = "CLIENTE"
   frmPessoaConsulta.Show 1

   If Trim(CNPJCPF_A) <> "" Then
      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Text = ""
      txtCNPJCPF.Mask = "##############"

      txtCNPJCPF.Text = CNPJCPF_A
      Call txtCNPJCPF_LostFocus
      'txtCNPJCPF.PromptInclude = True

      txtProduto.Enabled = True

      txtProduto.SetFocus
      Exit Sub
   End If
   CNPJCPF_A = ""
   txtCNPJCPF.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsCli_Click"
End Sub

Private Sub txtAltura_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtAltura.Text) = "" Then _
         Exit Sub
      txtLargura.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtAltura_KeyPress"
End Sub

Private Sub txtlargura_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtLargura.Text) = "" Then _
         Exit Sub
      txtQtde.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtLargura_KeyPress"
End Sub

Private Sub txtQTDE_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtQtde.Text) = "" Then _
         Exit Sub
      txtVlrUnitario.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtQtde_KeyPress"
End Sub

Private Sub txtVlrUnitario_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      If Trim(txtVlrUnitario.Text) = "" Then _
         Exit Sub

PROCESSA_ITEM

      txtProduto.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtVlrUnitario_KeyPress"
End Sub

Private Sub txtCNPJCPF_GotFocus()
'On Error GoTo ERRO_TRATA

   txtCNPJCPF.PromptInclude = False
   SQL3 = "" & txtCNPJCPF.Text
      txtCNPJCPF.Mask = "##############"
      txtCNPJCPF.Mask = ""
   txtCNPJCPF.Text = SQL3

   txtCNPJCPF.SelStart = 0
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.SelLength = Len(txtCNPJCPF)

   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTCNPJCPF_GotFocus"
End Sub

Private Sub txtCNPJCPF_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF2
         txtProduto.Enabled = True

         txtProduto.SetFocus
         SQL3 = txtNome.Text
         txtNome.Text = Trim(InputBox("Informe Nome do cliente", "Emissão de Cupom Fiscal", SQL3))
      Case vbKeyF7
         CNPJCPF_A = ""
         frmPessoaConsulta.Show 1
         If Trim(CNPJCPF_A) <> "" Then
            txtCNPJCPF.PromptInclude = False
            txtCNPJCPF.Text = CNPJCPF_A
         End If
      Case vbKeyBack
         If Not IsNumeric(txtCNPJCPF.Text) Then _
            txtCNPJCPF.Mask = "##############"
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyDown"
End Sub

Private Sub txtCNPJCPF_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtCNPJCPF.PromptInclude = False
      txtCNPJCPF.Mask = ""
      If Trim(txtCNPJCPF.Text) = "99999999999" Then
         txtNome.Enabled = True

         'txtProduto.Enabled = True
         'txtProduto.SetFocus
         Else
            txtProduto.Enabled = True
            'txtProduto.SetFocus
            'txtNome.Enabled = False
      End If
      txtNome.Enabled = True
      txtCNPJCPF.PromptInclude = False
      txtNome.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCNPJCPF_KeyPress"
End Sub

Private Sub txtCNPJCPF_LostFocus()
'On Error GoTo ERRO_TRATA

   'txtNome.Text = ""
   PESSOA_ID_N = 0
   CLIENTE_ID_N = 0
   CNPJCPF_CLIENTE_A = ""
   NOME_CLIENTE_A = ""
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) = "" Then
      txtCNPJCPF.Text = "99999999999"
      txtNome.Enabled = True
      If Trim(txtNome.Text) = "" Then _
         txtNome.Text = "Consumidor Final"
   End If
   txtCNPJCPF.PromptInclude = False
   If TRATA_CLIENTE(txtCNPJCPF.Text) = False Then
      CHECA_CLIENTE txtCNPJCPF.Text, txtNome.Text
      txtCNPJCPF.Text = "99999999999"
      txtNome.Enabled = True
      If Trim(txtNome.Text) = "" Then _
         txtNome.Text = "Consumidor Final"
      Else
         txtNome.Text = "" & NOME_CLIENTE_A
         'If Trim(txtNome.Text) = "" Then _
            txtNome.Text = "" & NOME_CLIENTE_A
   End If

   'txtPAGAR.Text = Format(VALOR_PENDENTE_N, strFormatacao2Digitos)
   'txtPAGAR.Refresh

   CRITERIO_A = txtCNPJCPF.Text
   
   If txtCNPJCPF.Text <> "" Then
      CRITERIO_A = txtCNPJCPF.Text

      If Not IsNull(txtCNPJCPF.Text) Then
          If Len(txtCNPJCPF.Text) <= 11 Then
              txtCNPJCPF.Mask = "###.###.###-##"
              Else
                If Len(txtCNPJCPF.Text) > 11 Then _
                    txtCNPJCPF.Mask = "##.###.###/####-##"
          End If
      End If
      txtCNPJCPF.Text = CRITERIO_A
   End If

   txtCNPJCPF.BackColor = &HFFFFFF
   txtCNPJCPF.PromptInclude = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtCnpjCpf_LostFocus"
End Sub

Private Sub txtNome_GotFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNome.Text) <> "" Then
      txtNome.SelStart = 0
      txtNome.SelLength = Len(txtNome)
   End If
   txtNome.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_GotFocus"
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      txtProduto.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNome_KeyPress"
End Sub

Private Sub cmdConsProd_Click()
'On Error GoTo ERRO_TRATA

   CONSULTA_PRODUTO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdConsProd_Click"
End Sub

Private Sub TXTPRODUTO_GotFocus()
'On Error GoTo ERRO_TRATA

   txtDescricao.Enabled = False
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
      Case vbKeyF6
         'If Trim(txtPedido.Text) <> "" And Trim(txtProduto.Text) <> "" And Trim(txtSeq.Text) <> "" Then _
            EXCLUIR_ITEM Trim(txtProduto.Text), Trim(txtPedido.Text), Trim(txtSeq.Text)

         txtProduto.Enabled = True
         txtProduto.SetFocus
      Case vbKeyF7
         CONSULTA_PRODUTO
         txtProduto.Enabled = True
         txtProduto.SetFocus
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
      If Trim(txtProduto.Text) = "" Then _
         Exit Sub

      PROCESSA_DADOS_PRODUTOS
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtproduto_KeyPress"
End Sub

Private Sub TXTPRODUTO_LostFocus()
   txtProduto.BackColor = &HFFFFFF
End Sub

Private Sub txtAltura_GotFocus()
   txtAltura.SelStart = 0
   txtAltura.SelLength = Len(txtAltura)
   txtAltura.BackColor = &HC0FFFF
End Sub

Private Sub txtlargura_GotFocus()
   txtLargura.SelStart = 0
   txtLargura.SelLength = Len(txtLargura)
   txtLargura.BackColor = &HC0FFFF
End Sub

Private Sub txtQTDE_GotFocus()
   txtQtde.SelStart = 0
   txtQtde.SelLength = Len(txtQtde)
   txtQtde.BackColor = &HC0FFFF
End Sub

Private Sub txtTotItem_GotFocus()
   txtTotItem.SelStart = 0
   txtTotItem.SelLength = Len(txtTotItem)
   txtTotItem.BackColor = &HC0FFFF
End Sub

Private Sub txtVlrUnitario_GotFocus()
   txtVlrUnitario.SelLength = Len(txtVlrUnitario)
   txtVlrUnitario.BackColor = &HC0FFFF
End Sub

Private Sub txtAltura_LostFocus()
   txtAltura.BackColor = &HFFFFFF
   If Trim(txtAltura.Text) <> "" Then _
      txtAltura.Text = Format(txtAltura.Text, strFormatacao3Digitos)
End Sub

Private Sub txtlargura_LostFocus()
   If Trim(txtLargura.Text) <> "" Then _
      txtLargura.Text = Format(txtLargura.Text, strFormatacao3Digitos)
   txtLargura.BackColor = &HFFFFFF
End Sub

Private Sub txtQTDE_LostFocus()
   If Trim(txtQtde.Text) <> "" Then _
      txtQtde.Text = Format(txtQtde.Text, strFormatacao3Digitos)
   txtQtde.BackColor = &HFFFFFF
   CALCULO_METRO_QUADRADA
End Sub

Private Sub txtVlrUnitario_LostFocus()
   If Trim(txtVlrUnitario.Text) <> "" Then _
      txtVlrUnitario.Text = Format(txtVlrUnitario.Text, strFormatacao3Digitos)
   txtVlrUnitario.BackColor = &HFFFFFF
End Sub

Private Sub txtTotItem_LostFocus()
   txtTotItem.BackColor = &HFFFFFF
End Sub

'=================
Sub INICIA_TELA()
'On Error GoTo ERRO_TRATA

   LIMPA_ORÇAMENTO
   MOSTRA_VENDEDORES
   MOSTRA_ESTAGIO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INICIA_TELA"
End Sub

Sub LIMPA_ORÇAMENTO()
'On Error GoTo ERRO_TRATA

   txtDtEmis.PromptInclude = False
   txtDtEmis.Text = Date
   txtDtEmis.PromptInclude = True
   txtPedido.Text = ""
   lblSituacao.Caption = ""
   cmbVendedorAUX.Text = ""
   cmbVendedor.Text = ""
   cmdEstagioAUX.Text = ""
   cmdEstagio.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   lblTotal.Caption = ""
   txtValorDig.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_ORÇAMENTO"
End Sub

Private Sub MOSTRA_VENDEDORES()
'On Error GoTo ERRO_TRATA

   cmbVendedor.Clear
   cmbVendedorAUX.Clear

   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   SQL = "select descricao,vendedor_id from vwVendedor WITH (NOLOCK)"
   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabVENDEDOR.EOF
      cmbVendedor.AddItem Trim(TabVENDEDOR!DESCRICAO) & "-" & Trim(TabVENDEDOR!VENDEDOR_ID)
      cmbVendedorAUX.AddItem Trim(TabVENDEDOR!VENDEDOR_ID)
      TabVENDEDOR.MoveNext
   Wend
   If TabVENDEDOR.State = 1 Then _
      TabVENDEDOR.Close

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select logon from USUARIO WITH (NOLOCK)"
   SQL = SQL & " where usuario_id = " & USUARIO_ID_N
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then
      cmbVendedor.Enabled = False

      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close

      CRITERIO_A = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
      SQL = "select descricao, vendedor_id from vwVendedor WITH (NOLOCK)"
      SQL = SQL & " where descricao like " & CRITERIO_A
      TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabVENDEDOR.EOF Then
         cmbVendedor.Text = Trim(TabVENDEDOR!DESCRICAO)
         cmbVendedorAUX.Text = Trim(TabVENDEDOR!VENDEDOR_ID)
      End If
      If TabVENDEDOR.State = 1 Then _
         TabVENDEDOR.Close
   End If

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then _
      cmbVendedor.Enabled = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_VENDEDORES"
End Sub

Sub MOSTRA_ESTAGIO()
'On Error GoTo ERRO_TRATA

   cmdEstagio.Clear
   cmdEstagioAUX.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from DESCR "
   SQL = SQL & " where TIPO = 'P1'"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      cmdEstagio.AddItem Trim(TabConsulta!DESCRICAO)
      cmdEstagioAUX.AddItem Trim(TabConsulta!codigo)
      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_ESTAGIO"
End Sub

Sub MOSTRA_TURNO()
'On Error GoTo ERRO_TRATA

   adoTurno.ConnectionString = AUTENTICA_GRID
   adoTurno.CommandType = adCmdText

   SQL = "select * from TURNO WITH (NOLOCK) "

   adoTurno.RecordSource = SQL
   adoTurno.Enabled = True
   adoTurno.Refresh
   grdTurno.Refresh

   grdTurno.Columns(0).DataField = "TURNO_ID"
   grdTurno.Columns(0).Caption = "Turno"
   grdTurno.Columns(0).Width = 800
   grdTurno.Columns(0).Alignment = dbgLeft

   grdTurno.Columns(1).DataField = "HoraIni"
   grdTurno.Columns(1).Caption = "HoraIni"
   grdTurno.Columns(1).Width = 1100
   grdTurno.Columns(1).Alignment = dbgLeft

   grdTurno.Columns(2).DataField = "HoraFim"
   grdTurno.Columns(2).Caption = "HoraFim"
   grdTurno.Columns(2).Width = 1100
   grdTurno.Columns(2).Alignment = dbgLeft

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_TURNO"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtPedido.Text = ""
   lblSituacao.Caption = ""
   cmbVendedorAUX.Text = ""
   cmbVendedor.Text = ""
   cmdEstagioAUX.Text = ""
   cmdEstagio.Text = ""
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = ""
   txtNome.Text = ""
   lblTotal.Caption = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub PROCESSA_DADOS_PRODUTOS()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = 0
   If Trim(txtPedido.Text) <> "" Then _
      If IsNumeric(txtPedido.Text) Then _
         PEDIDO_ID_N = 0 & Trim(txtPedido.Text)

   If (LE_PRODUTO(Trim(txtProduto.Text), "C")) = False Then
      txtProduto.Enabled = True
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto)
      Exit Sub
   End If

   If Trim(UNIDADE_MEDIDA_A) <> "" Then
      If Trim(UNIDADE_MEDIDA_A) = "M²" Then
         txtAltura.Visible = True
         txtLargura.Visible = True
         lblAltura.Visible = True
         lblLargura.Visible = True

         txtAltura.SetFocus
         Else
            txtUN.Text = "" & Trim(UNIDADE_MEDIDA_A)
            txtAltura.Visible = False
            txtLargura.Visible = False
            lblAltura.Visible = False
            lblLargura.Visible = False

            txtQtde.SetFocus
      End If
   End If

   txtDescricao.Text = "" & Trim(DESC_PRODUTO_A)
   txtQtde.Text = Format(QTDE_N, strFormatacao3Digitos)
   txtVlrUnitario.Text = "" & Format(PR_VAREJO_N, strFormatacao3Digitos)
   'txtVarejo.Text = "" & Format(PR_VAREJO_N, strFormatacao2Digitos)
   'txtAtacado.Text = "" & Format(PR_ATACADO_N, strFormatacao2Digitos)
   txtQtde.Text = Format(QTDE_N, strFormatacao3Digitos)
   'txtPreçoCusto.Text = "" & Format(PR_CUSTO_PRODUTO_N, strFormatacao2Digitos)

   If INDR_ESTQ_NEGATIVO = False Then
      QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

      If QTDE_ESTOQUE_N <= 0 Then
         MsgBox "Produto sem estoque disponível."
         txtProduto.Enabled = True
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)
         txtProduto.SetFocus
         Exit Sub
      End If
   End If
   If Len(CODG_NCM_A) > 2 Then
      If Len(CODG_NCM_A) < 8 Then
         MsgBox "Cadastro do produto : " & Trim(txtDescricao.Text) & " está incorreto, verificar código NCM !!!"

         LIMPA_BODY

         txtProduto.Enabled = True
         txtProduto.SelStart = 0
         txtProduto.SelLength = Len(txtProduto)

         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If STATUS_PROD = "P" Then
      txtProduto.ForeColor = vbRed
      txtDescricao.ForeColor = vbRed
      Else
         If STATUS_PROD = "C" Then
            MsgBox "Produto desativado para venda , Favor Confirmar!"
            txtProduto.Enabled = True
            txtProduto.SelStart = 0
            txtProduto.SelLength = Len(txtProduto)
            txtProduto.SetFocus
            Exit Sub
         End If
   End If
'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
   txtSeq.Text = SEQ_ID_N
'=====================
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select * from PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and pedido_ID = " & PEDIDO_ID_N
   SQL = SQL & " and seq_ID = " & Trim(txtSeq.Text)
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      txtVlrUnitario.Text = "" & Format(TabPedidoItem!Valor_Item, strFormatacao2Digitos)
      QTDE_PEDIDO = 0 & TabPedidoItem!QTD_PEDIDA
      txtQtde.Text = "" & Format(QTDE_PEDIDO, strFormatacao3Digitos)
      VALOR_ITEM_N = 0 & TabPedidoItem!Valor_Item
      VALOR_DIFERENCA_N = 0 & (TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA)
      txtSeq.Text = "" & TabPedidoItem.Fields("seq_id").Value
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   If INDR_LEU_POR_CODG_BARRAS = True Then
      txtQtde.Text = 1

      Call PROCESSA_ITEM

      CODIGO_BARRAS_A = ""
      txtProduto.Enabled = True
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto)
      txtProduto.SetFocus
      CODIGO_BARRAS_A = ""
      Exit Sub
   End If

   If Len(Trim(CODIGO_BARRAS_A)) = 13 Then
      If QTDE_N > 0 Then
         If Trim(txtVlrUnitario.Text) <> "" Then
            If IsNumeric(txtVlrUnitario.Text) Then

               Call PROCESSA_ITEM

               CODIGO_BARRAS_A = ""
               txtProduto.Enabled = True
               txtProduto.SelStart = 0
               txtProduto.SelLength = Len(txtProduto)
               txtProduto.SetFocus
               CODIGO_BARRAS_A = ""
               Exit Sub
            End If
         End If
      End If
      'Else: txtQtde.SetFocus
   End If
   CODIGO_BARRAS_A = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_DADOS_PRODUTOS"
End Sub

Sub PROCESSA_ITEM()
'On Error GoTo ERRO_TRATA

   If Trim(txtQtde.Text) <> "" Then
      If IsNumeric(txtQtde.Text) Then
         QTDE_N = txtQtde.Text
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

   txtCNPJCPF.PromptInclude = False
   If Trim(UF_CLIENTE_A) = "" Then _
      TRATA_CLIENTE txtCNPJCPF.Text

   If Trim(cmbVendedorAUX.Text) = "" Then
      cmbVendedor.Text = "BALCAO"
      cmbVendedorAUX.Text = 0
   End If

   txtCNPJCPF.PromptInclude = False
   If txtCNPJCPF.Text = "" Then _
      txtCNPJCPF.Text = "99999999999"

   If Trim(txtProduto.Text) = "" Then
      MsgBox "Informe codigo de Produto.", vbOKOnly, "Atenção."
      FraSeq.Enabled = True
      txtProduto.Enabled = True

      txtProduto.SetFocus
      Exit Sub
   End If

   If Not IsNull(txtVlrUnitario.Text) Then
      VALOR_ITEM_N = 0 & txtVlrUnitario.Text
      If VALOR_ITEM_N <= 0 Then
         MsgBox "Produto sem preço de venda.", vbOKOnly, "Atenção."
         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus
         Exit Sub
      End If
   End If

   If Trim(txtQtde.Text) = "" Then
      Beep
      MsgBox "Informe a quantidade.", vbOKOnly, "Atenção."
      txtQtde.SetFocus
      Exit Sub
      Else
         'quantidade pedida
         QTDE_PEDIDO = txtQtde.Text
         txtQtde.Text = Format(QTDE_PEDIDO, strFormatacao3Digitos)
         If INDR_CONTROLA_ESTOQUE = True Then
            CODG_PRODUTO_A = Trim(txtProduto.Text)

            If INDR_ESTQ_NEGATIVO = False Then
               QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

               If QTDE_ESTOQUE_N < 0 Then
                  Beep
                  MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
                  txtQtde.SetFocus
                  Exit Sub
               End If
            End If
         End If
         If QTDE_PEDIDO <= 0 Then
            Beep
            MsgBox "Quantidade pedida não permitido, deve ser maior que 0.", vbOKOnly, "Atenção."
            txtQtde.SetFocus
            Exit Sub
         End If
   End If

   'valor venda item
   VALOR_ITEM_N = txtVlrUnitario.Text
   VALOR_TOTAL_DESCONTO_N = 0

   'valor total da Pedido, o desconto é armazenado no seu devido lugar, não entra no calculo do campo total da venda
   VALOR_TOTAL_N = VALOR_TOTAL_N + (VALOR_ITEM_N * QTDE_PEDIDO) - VALOR_DIFERENCA_N

'===================
   'checa se o funcionário pode comprar produtos de produção conforme a cota diária estabelecida
   txtCNPJCPF.PromptInclude = False
   If Trim(txtCNPJCPF.Text) <> "99999999999" And Valor_Compra_Dia_Permitida > 0 And INDR_PRODUTO_PRODUCAO = True Then
      'If CHECA_FUNCIONARIO(txtCNPJCPF.Text) = True Then
      '   If CHECA_VALOR_DIARIO_PERMITIDO_PRODUCAO(ESTABELECIMENTO_ID_N, Date, Trim(txtCNPJCPF.Text), (QTDE_PEDIDO * VALOR_ITEM_N)) = True Then
      '      MsgBox "Cota diária de compra de produtos de produção ultrapassada, não permitido."
      '      Exit Sub
      '   End If
      'End If
   End If
'===================

   If Trim(txtPedido.Text) = "" Then
      GERA_PEDIDO_ID
      'ABRE_PEDIDO

      txtPedido.Text = PEDIDO_ID_N
   End If
   If Not IsNumeric(txtPedido.Text) Then
      GERA_PEDIDO_ID
      'ABRE_PEDIDO

      txtPedido.Text = PEDIDO_ID_N
   End If
   If Trim(txtPedido.Text) <> "" Then
      If IsNumeric(txtPedido.Text) Then
         txtPedido.Enabled = True
            PEDIDO_ID_N = txtPedido.Text
         txtPedido.Enabled = False
         'If INDR_PEDIDO_VALIDO = False Then _
            VALIDA_PEDIDO_ID PEDIDO_ID_N
      End If
   End If

   If INDR_PEDIDO_VALIDO = False Then _
      GRAVA_CABECA "R"

   If Trim(txtPedido.Text) <> "" Then _
      If IsNumeric(txtPedido.Text) Then _
         GRAVA_TUDO_ITEM

   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCESSA_ITEM"
End Sub

Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtUN.Text = ""
   txtProduto.Text = ""
   txtDescricao.Text = ""
   txtAltura.Text = ""
   txtLargura.Text = ""
   txtQtde.Text = ""
   txtVlrUnitario.Text = ""
   txtTotItem.Text = ""
   txtSeq.Text = ""
   txtValorDig.Visible = False

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub CONSULTA_PRODUTO()
'On Error GoTo ERRO_TRATA

   frmProdutoConsulta.Show 1
   If SQL3 <> "" Then
      txtProduto.Enabled = True
      txtProduto.Text = SQL3
      txtProduto.SetFocus
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Sub CALCULO_METRO_QUADRADA()
'On Error GoTo ERRO_TRATA

'O caixilho tem 2,10m de altura
'por 1,60m de largura,
'portanto uma área de vidro de 3,36m²,
'isto já sendo benevolente com o vidraceiro pois não
'estamos descontando a área dos montantes e das travessas internas.
'Telefonei para vários vidraceiros e os preços estavam em torno de R$ 60 o m², ou seja, ficaria por algo como R$ 201,60.

'p = (C + M + T) * E

'Onde:
'P = Preço de Venda
'C = Custo dos materiais empregados
'M = Custo da mão de obra
'T = Serviço de terceiros, se houver
'E = Encargos

'Resumindo, o nosso vidro sairia por:
'Material = 3,36 x 30 = 100,80
'Mão de obra = 2 horas a R$ 22 = 44,00
'Preço de venda = (100,80 + 44) * 1,43 = R$ 207,00

'[10:04, 7/3/2018] Leonardo CPS Portas: O PERFIL E METRO LINEAR
'[10:04, 7/3/2018] Leonardo CPS Portas: O VIDRO E QUADRADO
'[10:04, 7/3/2018] Leonardo CPS Portas: EXEMPLO
'[10:04, 7/3/2018] Leonardo CPS Portas: ESSE DV 3528
'[10:05, 7/3/2018] Leonardo CPS Portas: 1 METRO DELE CUSTA 17,00
'[10:05, 7/3/2018] Leonardo CPS Portas: 2 METRO 34,00
'[10:05, 7/3/2018] Leonardo CPS Portas: O VIDRO
'[10:05, 7/3/2018] Leonardo CPS Portas: O METRO QUADRADO CUSTA 115,00
'[10:05, 7/3/2018] Leonardo CPS Portas: 1X 1 = 115,00
'[10:05, 7/3/2018] Leonardo CPS Portas: 1,5 X 1 = 172,5
'[10:06, 7/3/2018] Leonardo CPS Portas: ENTENDEU ?
'AI O PERFIL E LINEAR E O VIDRO QUADRADO

   Dim ALTURA_N      As Double
   Dim LARGURA_N     As Double
   Dim QTDE_N        As Double
   Dim VALOR_METRO_QUADRADO_N As Double
   Dim VALOR_ITEM_N  As Double
   Dim VALOR_TOTAL_N As Double

   ALTURA_N = 0 & txtAltura.Text
   LARGURA_N = 0 & txtLargura.Text
   QTDE_N = 0 & txtQtde.Text

   VALOR_ITEM_N = 0 & PR_VAREJO_N

   If Trim(UNIDADE_MEDIDA_A) <> "" Then
      If Trim(UNIDADE_MEDIDA_A) = "M²" Then
         VALOR_METRO_QUADRADO_N = LARGURA_N * ALTURA_N
         VALOR_ITEM_N = 0 & (VALOR_METRO_QUADRADO_N * VALOR_ITEM_N)
         'Else: VALOR_ITEM_N = 0 & (VALOR_ITEM_N)
      End If
   End If

   txtVlrUnitario.Text = "" & Format(VALOR_ITEM_N, strFormatacao3Digitos)

   DoEvents

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CONSULTA_PRODUTO"
End Sub

Private Sub GRAVA_CABECA(TIPO_REGISTRO_A As String)
'On Error GoTo ERRO_TRATA

   CRITERIO_A = ""
   CLIENTE_ID_N = 0

   txtCNPJCPF.PromptInclude = False
   CHECA_CLIENTE txtCNPJCPF.Text, txtNome.Text

   PEDIDO_ID_N = 0 & txtPedido.Text

   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   SQL = "select status,nome_cliente,pedido_id from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & txtPedido.Text
   TabCABECA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCABECA.EOF Then
      If Not IsNull(TabCABECA.Fields("nome_cliente").Value) Then _
         txtNome.Text = Trim(TabCABECA.Fields("nome_cliente").Value)

      If IsNull(TabCABECA!Status) Then _
         Exit Sub

      'Emitido com Nota
      If Trim(TabCABECA.Fields("Status").Value) = 3 Then _
         Exit Sub

      If Trim(TabCABECA.Fields("Status").Value) = 4 Then _
         Exit Sub

      'Apenas Faturado
      If Trim(TabCABECA.Fields("Status").Value) = 5 Then _
         Exit Sub

      'Emitido com Cupom
      If Trim(TabCABECA.Fields("Status").Value) = 7 Then _
         Exit Sub

      'cancelado
      If Trim(TabCABECA.Fields("Status").Value) = 9 Then _
         Exit Sub

      PEDIDO_ID_N = 0 & TabCABECA.Fields("pedido_id").Value
      txtPedido.Text = PEDIDO_ID_N
      txtCNPJCPF.PromptInclude = False

      SQL = "UPDATE PEDIDO SET "
      SQL = SQL & " Valor_total = " & tpMOEDA(VALOR_TOTAL_N)
      SQL = SQL & ",pedido_id = " & txtPedido.Text
      SQL = SQL & ",Valor_desconto = 0"   'vai zerar e tratar somente na tela de desconto
      SQL = SQL & ",Perc_desc = " & tpMOEDA(0)
      SQL = SQL & ",CGCCPF = '" & Trim(txtCNPJCPF.Text) & "'"
      SQL = SQL & ",Vendedor_id = " & cmbVendedorAUX.Text
      SQL = SQL & ",dt_req = '" & Now & "'"
      SQL = SQL & ",nome_cliente = '" & txtNome.Text & "'"
      SQL = SQL & ",Status = 1 "
      SQL = SQL & ",TIPO_REGISTRO = '" & TIPO_REGISTRO_A & "'"
      SQL = SQL & ",usuario_id = " & USUARIO_ID_N
      SQL = SQL & ",TIPOvenda_id = " & 1 '& cmbFaturaAUX.Text
      SQL = SQL & ",USUARIO_LIBERA_VENDA = " & USUARIO_ID_N
      SQL = SQL & ",CLIENTE_ID = " & CLIENTE_ID_N
      SQL = SQL & ",cartaobarra_id = " & CARTAOBARRA_ID_N

      SQL = SQL & " where pedido_id = " & txtPedido.Text
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      Else
         If TabCABECA.State = 1 Then _
            TabCABECA.Close

         SQL = "INSERT INTO PEDIDO "
            SQL = SQL & "("
               SQL = SQL & "PEDIDO_ID,Empresa_id, CGCCPF, Vendedor_id, Dt_Req, Nome_Cliente, "
               SQL = SQL & " Tipo_Registro,usuario_id, TIPOVENDA_ID, CLIENTE_ID, Valor_ToTal,"
               SQL = SQL & " valor_desconto,perc_desc,NUMERO_CAIXA_CPU,ESTABELECIMENTO_ID,cartaobarra_id, Status"
            SQL = SQL & ")"
            SQL = SQL & " VALUES ("
               SQL = SQL & PEDIDO_ID_N
               SQL = SQL & "," & EMPRESA_ID_N
               SQL = SQL & ",'" & Trim(txtCNPJCPF.Text) & "'"
               SQL = SQL & "," & cmbVendedorAUX.Text & ","
               SQL = SQL & "'" & Now & "'"
               SQL = SQL & ",'" & Trim(txtNome.Text) & "'"
               SQL = SQL & ",'" & TIPO_REGISTRO_A & "'"
               SQL = SQL & "," & USUARIO_ID_N
               SQL = SQL & "," & 1      'cmbFaturaAux.Text
               SQL = SQL & "," & CLIENTE_ID_N
               SQL = SQL & "," & tpMOEDA(VALOR_TOTAL_N)
               SQL = SQL & "," & tpMOEDA(0)                       'vai zerar e tratar somente na tela de desconto
               SQL = SQL & "," & tpMOEDA(0)
               SQL = SQL & "," & NUMERO_CAIXA_CPU                 'NUMERO_CAIXA_CPU
               SQL = SQL & "," & ESTABELECIMENTO_ID_N
               SQL = SQL & "," & CARTAOBARRA_ID_N
               SQL = SQL & ",1"                                    'status
            SQL = SQL & ")"
   End If
   If TabCABECA.State = 1 Then _
      TabCABECA.Close

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_CABECA"
End Sub

Private Sub GRAVA_TUDO_ITEM()
'On Error GoTo ERRO_TRATA

   'Tratamento da tributacao
   'fazer no final desta rotina
   'CODG_PRODUTO_A = Trim(txtProduto.Text)

   If USA_NFe = True And INDR_CAIXA = False Then
      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
         If Trim(UF_CLIENTE_A) = "" Then
            If INDR_PEDIDO_VENDA = False Then
               MsgBox "Cliente com cadastro incompleto !!!"
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If

   'If Trim(txtPreçoCusto.Text) = "" Then _
      txtPreçoCusto.Text = 0

   'If Not IsNumeric(txtPreçoCusto.Text) Then _
      txtPreçoCusto.Text = 0

   If Trim(txtAltura.Text) = "" Then _
      txtAltura.Text = 0
   If Trim(txtLargura.Text) = "" Then _
      txtLargura.Text = 0

'=====================
   If Trim(txtSeq.Text) = "" Then
      SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
      Else
         If Not IsNumeric(txtSeq.Text) Then
            SEQ_ID_N = 0 & MAX_ID("seq_id", "PEDIDOITEM", "pedido_id", Trim(txtPedido.Text), "", "")
            Else: SEQ_ID_N = txtSeq.Text
         End If
   End If
'=====================
   Dim STATUS_ITEM_A As String
   STATUS_ITEM_A = ""

   STATUS_ITEM_A = "P"

   Dim TabPedidoItem       As New ADODB.Recordset

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "select pedido_id from PEDIDOITEM  WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and seq_id = " & SEQ_ID_N
   SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and status <> 'C' "

   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabPedidoItem.EOF Then
      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

      SQL = "INSERT INTO PEDIDOITEM "
      SQL = SQL & " ("
         SQL = SQL & "PEDIDO_ID,SEQ_ID,PRODUTO_ID, Qtd_Pedida,Valor_item, "
         SQL = SQL & " PERC_DESC, valor_desconto, status,preco_custo,TIPO_REG,"
         SQL = SQL & " PESO_ITEM,usu_atende,altura,largura"
      SQL = SQL & ") "
      SQL = SQL & " VALUES ("
         SQL = SQL & txtPedido.Text                                                       'PEDIDO_id
         SQL = SQL & "," & SEQ_ID_N                                                       'SEQ_ID
         SQL = SQL & "," & PRODUTO_ID_N                                                   'produto_id
         SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                                           'Qtd_Pedida
         SQL = SQL & "," & tpMOEDA(VALOR_ITEM_N)                                          'Valor_item
         SQL = SQL & "," & tpMOEDA(PERC_DESCONTO_N)                                       'PERC_DESC
         SQL = SQL & "," & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)  'valor_desconto
         SQL = SQL & ",'" & STATUS_ITEM_A & "'"                                           'status
         SQL = SQL & "," & tpMOEDA(PR_CUSTO_PRODUTO_N)                                    'PRECO_CUSTO
         SQL = SQL & ",'PC'"                                                              'TIPO_REG
         SQL = SQL & "," & tpMOEDA(QTDE_PEDIDO)                                           'PESO_ITEM

         If ATENDENTE_ID_N > 0 Then
            SQL = SQL & "," & ATENDENTE_ID_N                                              'USU_ATENDE
            Else
               If USU_ATENDE_N <= 0 Then _
                  USU_ATENDE_N = USUARIO_ID_N
               SQL = SQL & "," & USU_ATENDE_N                                             'USU_ATENDE
         End If

         SQL = SQL & "," & tpMOEDA(txtAltura.Text)                                        'altura
         SQL = SQL & "," & tpMOEDA(txtLargura.Text)                                       'largura
      SQL = SQL & ")"
      Else
         If TabPedidoItem.State = 1 Then _
            TabPedidoItem.Close

         SQL = "UPDATE PEDIDOITEM SET "
         SQL = SQL & " qtd_pedida = " & tpMOEDA(QTDE_PEDIDO)
         SQL = SQL & ", Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
         SQL = SQL & ", PERC_desc = " & tpMOEDA(PERC_DESCONTO_N)
         SQL = SQL & ", valor_desconto = " & tpMOEDA((VALOR_ITEM_N * QTDE_PEDIDO) * PERC_DESCONTO_N / 100)
         SQL = SQL & ", status = '" & STATUS_ITEM_A & "'"
         SQL = SQL & ", preco_custo = " & tpMOEDA(PR_CUSTO_PRODUTO_N)
         SQL = SQL & ", PESO_ITEM = " & tpMOEDA(QTDE_PEDIDO)
         SQL = SQL & ", altura = " & tpMOEDA(txtAltura.Text)
         SQL = SQL & ", largura = " & tpMOEDA(txtLargura.Text)

         SQL = SQL & " Where pedido_id = " & txtPedido.Text
         SQL = SQL & " and seq_id = " & SEQ_ID_N
   End If
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   CONECTA_RETAGUARDA.Execute SQL
   INDR_PEDIDO_VALIDO = True

   'Tratamento da tributacao
   CODG_PRODUTO_A = Trim(txtProduto.Text)
   USU_ATENDE_N = 0

   txtCNPJCPF.PromptInclude = False
   PREPARA_TRIBUTACAO_PRODUTO

   SETA_GRID

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_TUDO_ITEM"
End Sub

Private Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) = "" Then _
      Exit Sub
   If Not IsNumeric(txtPedido.Text) Then _
      Exit Sub
   If PEDIDO_ID_N <= 0 Then _
      Exit Sub

   Dim TabGrid As New ADODB.Recordset
   Dim Coluna, Linha, Largura_Campo
   Dim VALOR_ITENS_PRODUCAO   As Double
   Dim VALOR_ITENS_REVENDA    As Double

   CONT_N = 0
   VALOR_DESCONTO_N = 0
   VALOR_ITEM_N = 0
   VALOR_TOTAL_N = 0
   VALOR_ITENS_PRODUCAO = 0
   VALOR_ITENS_REVENDA = 0

   'txtItens.Text = "" & CONT_N
   lblTotal.Caption = "Total Orçamento = " & Format(VALOR_TOTAL_N, "currency")

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

   SQL = "select * from vwPedidoVendaItens"

   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and StatusItem <> 'C' "
   SQL = SQL & " order by seq_id desc"

   TabGrid.Open SQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGrid.EOF Then
      INDR_PEDIDO_VALIDO = True
      'txtCNPJCPF.Enabled = False
      'cmdConsCli.Enabled = False

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
         If Not IsNull(TabGrid.Fields("producao").Value) Then _
            If TabGrid.Fields("producao").Value = True Then _
               INDR_PRI = True

'=======totais
         CONT_N = CONT_N + 1
         VALOR_ITEM_N = VALOR_ITEM_N + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtde").Value)
         If Not IsNull(TabGrid.Fields("desconto").Value) Then _
            VALOR_DESCONTO_N = VALOR_DESCONTO_N + TabGrid.Fields("desconto").Value

         If INDR_PRI = True Then
            'VALOR_ITENS_PRODUCAO = 0
            VALOR_ITENS_PRODUCAO = VALOR_ITENS_PRODUCAO + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtde").Value)
            Else
               'VALOR_ITENS_REVENDA = 0
               VALOR_ITENS_REVENDA = VALOR_ITENS_REVENDA + (TabGrid.Fields("valoritem").Value * TabGrid.Fields("qtde").Value)
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

'Codigo Produto
      MSFlexGrid1.ColWidth(0) = 2000
      MSFlexGrid1.ColAlignment(0) = 0

'Referencia
      MSFlexGrid1.ColWidth(1) = 0
      MSFlexGrid1.ColAlignment(1) = 0

'Descrição Produto
      MSFlexGrid1.ColWidth(2) = 7000
      MSFlexGrid1.ColAlignment(2) = 0

'QTDE
      MSFlexGrid1.ColWidth(3) = 2000
      MSFlexGrid1.ColAlignment(3) = 7

'Valor Item
      MSFlexGrid1.ColWidth(4) = 2000
      MSFlexGrid1.ColAlignment(4) = 7

'Desconto
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(5) = 0
         Else: MSFlexGrid1.ColWidth(5) = 1500
      End If
      MSFlexGrid1.ColAlignment(5) = 7

'Total Item
      MSFlexGrid1.ColWidth(6) = 2000
      MSFlexGrid1.ColAlignment(6) = 7

'SITUAÇÃO TRIBUTARIA PRODUTO
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(7) = 1
         Else: MSFlexGrid1.ColWidth(7) = 500
      End If
      MSFlexGrid1.ColAlignment(7) = 0

'ALIQUOTA ICMS
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(8) = 1
         Else: MSFlexGrid1.ColWidth(8) = 500
      End If
      MSFlexGrid1.ColAlignment(8) = 0

'NCM
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(9) = 1
         Else: MSFlexGrid1.ColWidth(9) = 500
      End If
      MSFlexGrid1.ColAlignment(9) = 0

'Pedido_id
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(10) = 1
         Else: MSFlexGrid1.ColWidth(10) = 500
      End If
      MSFlexGrid1.ColAlignment(10) = 0

'seq_id
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(11) = 1
         Else: MSFlexGrid1.ColWidth(11) = 500
      End If
      MSFlexGrid1.ColAlignment(11) = 0

'produto_id
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(12) = 1
         Else: MSFlexGrid1.ColWidth(12) = 500
      End If
      MSFlexGrid1.ColAlignment(12) = 0

'SITUAÇÃO ITEM
      If MULT_EMPRESA_B = True Then
         MSFlexGrid1.ColWidth(13) = 1
         Else: MSFlexGrid1.ColWidth(13) = 500
      End If
      MSFlexGrid1.ColAlignment(13) = 0

'familiaproduto_id
      MSFlexGrid1.ColWidth(14) = 50
      MSFlexGrid1.ColAlignment(14) = 0

'producao
      MSFlexGrid1.ColWidth(15) = 0
      MSFlexGrid1.ColAlignment(15) = 0
   End If

   ' fecha o recordset e a conexao
   If TabGrid.State = 1 Then _
      TabGrid.Close

   'txtItens.Text = "" & CONT_N

   VALOR_TOTAL_N = VALOR_ITEM_N - VALOR_DESCONTO_N

   lblTotal.Caption = "Total Orçamento = " & Format(VALOR_TOTAL_N, "currency")
   DoEvents

MSFlexGrid1.Visible = True
   VALOR_ITENS_PRODUCAO = 0
   VALOR_ITENS_REVENDA = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Private Sub PREPARA_TRIBUTACAO_PRODUTO()
'On Error GoTo ERRO_TRATA

'UF DO CLIENTE
'IVA
'MVA (Margem de Valor Agregado)
'PEGAR ALIQUOTA PARA O ESTADO TAL

   If Trim(CODG_PRODUTO_A) = "" Then
      MsgBox "Produto não informado, verifique !!!"
      Exit Sub
   End If

   If Trim(UF_CLIENTE_A) = "" Then _
      TRATA_CLIENTE txtCNPJCPF.Text

'MsgBox "PREPARA_TRIBUTACAO_PRODUTO  ====  " & CLIENTE_ID_N

   If Trim(UF_EMPRESA_A) = "" Then _
      PEGA_DADOS_EMPRESA

   If CLIENTE_ID_N < 0 Then
      MsgBox "Cliente não informado, verifique !!!"
      Exit Sub
   End If

   Dim rstProduto                As New ADODB.Recordset
   Dim RstTemp                   As New ADODB.Recordset
   Dim ST_PRODUTO_A              As String
   Dim VALOR_BASE_ICMS_N         As Double
   Dim VALOR_PERC_ICMS_N         As Double
   Dim VALOR_ICMS_PRODUTO        As Double
   Dim VALOR_BASE_ICMS_SUBST_N   As Double
   Dim VALOR_ICMS_PRODUTO_SUBST_N   As Double
   Dim VALOR_PERC_ICMS_SUBST_N   As Double
   Dim strCFOP_ITEM              As String
   Dim PERC_REDUCAO_ICMS_N       As Double
   Dim PERC_IVA_N                As Double
   Dim VALOR_TOTAL_ITEM_N        As Double
   Dim Aliquota_N                As Double

   VALOR_BASE_ICMS_N = 0
   VALOR_PERC_ICMS_N = 0
   VALOR_ICMS_PRODUTO = 0
   VALOR_BASE_ICMS_SUBST_N = 0
   VALOR_ICMS_PRODUTO_SUBST_N = 0
   VALOR_PERC_ICMS_SUBST_N = 0
   PERC_REDUCAO_ICMS_N = 0
   PERC_IVA_N = 0
   Aliquota_N = 0
   VALOR_TOTAL_ITEM_N = 0

   strCFOP_ITEM = ""
   strCFOP_ITEM = "5102"

   ST_PRODUTO_A = ""
   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0

   If USA_NFe = True And INDR_CAIXA = False Then
      txtCNPJCPF.PromptInclude = False
      If Trim(txtCNPJCPF.Text) <> "99999999999" And Trim(txtCNPJCPF.Text) <> "" Then
         If Trim(UF_CLIENTE_A) = "" Then
            If INDR_PEDIDO_VENDA = False Then
               MsgBox "Cliente com cadastro incompleto !!!"
               txtCNPJCPF.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If

   VALOR_ITEM_N = 0 & txtVlrUnitario.Text
   QTDE_N = 0 & txtQtde.Text
   VALOR_TOTAL_ITEM_N = (QTDE_N * VALOR_ITEM_N)

   If rstProduto.State = 1 Then _
      rstProduto.Close

   SQL = "select situacao_tributaria,perciva,comp_tributaria from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   SQL = SQL & " and situacao <> 'C' "
   rstProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstProduto.EOF Then
      If rstProduto.State = 1 Then _
         rstProduto.Close

      MsgBox "O sistema nao localizou nenhum produto com o seguinte codigo: " & CODG_PRODUTO_A & vbCrLf & "Verique"
      Exit Sub
   End If

ST_PRODUTO_A = "" & rstProduto!SITUACAO_TRIBUTARIA

   'Inicio yuri 01/05/2012
   ' Aqui será colocado a rotina para calcular os tributos em substituição a toda essa regra que esta
   ' nesta instrução
   ' busca aliquota do Unidade federativa do Cliente
   ' aqui nao retirar aqui vamos dar o inicio a toda carga tributaria
   ' comentei aqui para nao atraplhar se codigo

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0

'set parei aqui
Call BUSCA_ALIQUOTA_ICMS(UF_CLIENTE_A, "")

   ' fim yuri 01/05/2012, HORACIO MEXEU EM 06/09/2016

'ST_PRODUTO_A  = vem do cadastro de produto

   'Tributada integralmente
   If Trim(ST_PRODUTO_A) = "00" Then
      '5405  Venda de mercadoria, adquirida ou recebida de terceiros,
      'sujeita ao regime de substituição tributária,
      'na condição de contribuinte-substituído

      'Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros
      'em operação com mercadorias sujeitas ao regime de substituição tributária,
      'na condição de contribuinte substituído.
      strCFOP_ITEM = "5405"   'não é industria

      'se é optante do simples nacional
      If CTR_EMPRESA_N = 1 Then
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
            strCFOP_ITEM = "5102"
            Else: strCFOP_ITEM = "6102"
         End If
      End If

      If INDR_INDUSTRIA = True Then
         'Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros,
         'na condição de contribuinte substituto,
         'em operação com mercadorias sujeitas ao regime de substituição tributária.
         strCFOP_ITEM = "5403"

         'se é optante do simples nacional
         If CTR_EMPRESA_N = 1 Then
            If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
'INDR_PRODUTO_PRODUCAO = vem da tabela de familia informando que o produto é de produção
               If INDR_PRODUTO_PRODUCAO = True Then _
                  strCFOP_ITEM = "5101"
               Else
                  strCFOP_ITEM = "6102"
                  If INDR_PRODUTO_PRODUCAO = True Then _
                     strCFOP_ITEM = "6101"
            End If
         End If
      End If

      'Desconto nao entra no valor do ICMS de acordo com informacoes da CONTABILIDADE
      VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N

      'Criar campo de TIPO DE CLIENTE NO CADASTRO DE CLIENTE, não, mudou, se tem inscrição estadual é 2
      If TIPO_CLIENTE_N = 2 Then
'ESTAVA ASSIM
         'DENTRO DO ESTADO
         'If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
         '   VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CONTRIB) / 100) 'Valor da Reducao da base
         '   VALOR_PERC_ICMS_N = TP2_DE_CONTRIB                                'Percentual da reducao
         'End If

'==================EU
         'DENTRO DO ESTADO ICMS NORMAL
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
            VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * ALIQUOTA_ICMS_NORMAL_DENTRO_UF) / 100)  'Valor da Reducao da base
            VALOR_PERC_ICMS_N = ALIQUOTA_ICMS_NORMAL_DENTRO_UF                                 'Percentual da reducao
         End If
         'FORA DO ESTADO ICMS NORMAL
         If Trim(UF_CLIENTE_A) <> Trim(UF_EMPRESA_A) Then
            VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * ALIQUOTA_ICMS_NORMAL_FORA_UF) / 100)  'Valor da Reducao da base
            VALOR_PERC_ICMS_N = ALIQUOTA_ICMS_NORMAL_FORA_UF                                 'Percentual da reducao
         End If
      End If
   End If

'Margem de Valor Ajustada – MVA

   'Tributada e com cobrança do ICMS por substituição tributária
   If Trim(ST_PRODUTO_A) = 10 Then 'Substituicao Tributaria
      VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N

      Aliquota_N = ALIQUOTA_ICMS_NORMAL_DENTRO_UF
      If Trim(UF_CLIENTE_A) <> Trim(UF_EMPRESA_A) Then _
         Aliquota_N = ALIQUOTA_ICMS_NORMAL_FORA_UF

      If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
         'Campo IVA nao existe nao tabela verificar se precisa, Índices de Valor Agregado
         If Not IsNull(rstProduto!PERCIVA) Then _
           VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * rstProduto!PERCIVA) / 100)  'Valor da Reducao da base

         'VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * 1) / 100)  'Valor da Reducao da base
         VALOR_ICMS_PRODUTO_SUBST_N = ((VALOR_BASE_ICMS_SUBST_N * Aliquota_N) / 100)  'é fixo o percentual, procurar saber se tem como parametrizar
         VALOR_PERC_ICMS_SUBST_N = Aliquota_N
      End If
   End If

   'Com redução de base de cálculo
   If Trim(ST_PRODUTO_A) = 20 Then 'Reducao da base de calculo
      If rstProduto!COMP_TRIBUTARIA = 0 Then 'tipos de maquinas, normais, agricolas, industriais
         If CCE_CLIENTE_A <> "" Then    'Tem que ter inscricao estadual
            VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CONTRIB) / 100)
            PERC_REDUCAO_ICMS_N = TP2_DE_CONTRIB
            Else  'Sem inscricao estadual
               VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NCONTRIB) / 100)
               PERC_REDUCAO_ICMS_N = TP2_DE_NCONTRIB
         End If
      End If

      'Maquinas agricolas
      If rstProduto!COMP_TRIBUTARIA = 1 Then
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
            If CCE_CLIENTE_A <> "" Then
               VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CMAQ_IMP) / 100)
               PERC_REDUCAO_ICMS_N = TP2_DE_CMAQ_IMP
               Else
                  VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NMAQ_IMP) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_NMAQ_IMP
            End If
            Else 'Fora do Estado
               If CCE_CLIENTE_A <> "" Then
                  VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_CMAQ_IMP) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_FE_CMAQ_IMP
                  Else
                     VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_NMAQ_IMP) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_FE_NMAQ_IMP
               End If
         End If
      End If

      If rstProduto!COMP_TRIBUTARIA = 2 Then 'Maquinas industriais
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
            If CCE_CLIENTE_A <> "" Then
               VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CONTRIB) / 100)
               PERC_REDUCAO_ICMS_N = TP2_DE_CONTRIB
               Else
                  VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NCONTRIB) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_NCONTRIB
            End If
            Else 'Fora do Estado
               If CCE_CLIENTE_A <> "" Then
                  VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_CAP_INDU) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_FE_CAP_INDU
                  Else
                     VALOR_BASE_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_NAP_INDU) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_FE_NAP_INDU
               End If
         End If
      End If
   End If

   'Isenta ou não tributada e com cobrança do ICMS por substituição tributária
   If Trim(ST_PRODUTO_A) = 30 Then '//Isenta ou nao Tributada Com ICMS por Subs. Trib
      VALOR_BASE_ICMS_N = 0
      VALOR_PERC_ICMS_N = 0

      If UCase(UF_CLIENTE_A) <> UCase(UF_EMPRESA_A) Then
          '//Desconto nao entra no valor de ICMS de Acordo com as
          '//Informacoes Contabeis
          '//move (ITENS.TOTAL_ITEM - ITENS.VLR_DESC_RATEIO)  ;
          '//                                     To   ITENS.VLR_BASE_ICMS
          VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
          '??? nao grava o percentual do aliquota?
      End If
   End If

   'Isenta ou Não tributada
   If Trim(ST_PRODUTO_A) = 40 Or Trim(ST_PRODUTO_A) = 41 Then '//Isento ou nao Tributado
      VALOR_BASE_ICMS_N = 0
      VALOR_PERC_ICMS_N = 0
   End If

'50      Suspensão
'51      Diferimento

   'ICMS cobrado anteriormente por substituição tributária
   If Trim(ST_PRODUTO_A) = 60 Then '//Situacao Tributaria com Substitui‡ao Tributaria
      '//Desconto nao entra no valor de ICMS de Acordo com as
      '//Informacoes Contabeis

      VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
      If UCase(UF_CLIENTE_A) = UCase(UF_EMPRESA_A) Then
         If TIPO_CLIENTE_N = 2 Then 'Atacado
            '//Dentro do Estado e Cliente Contribuinte ele e Isento
            '/Emanoel Informacoes Contabilidade dia 30/05/2006
            VALOR_BASE_ICMS_N = 0
            VALOR_PERC_ICMS_N = 0
         End If
         'Só é tratado o tipo de cliente 2, atacado, e os outros tipos de clientes (varejo),
         'nao precisa tratar?
         Else 'Fora do estado
            If TIPO_CLIENTE_N = 2 Then 'Atacado
               VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
               'nao grava o percentual? porque?
            End If
      End If
   End If

'70      Com redução de base de cálculo e cobrança de ICMS por substituição tributária
'90      Outras

'========================================================================
'========================================================================
'========================================================================

   'If Not IsNull(rstProduto.Fields("CFOP_id").Value) Then
      
   'End If

   'DENTRO DO ESTADO
   If UCase(UF_CLIENTE_A) = UCase(UF_EMPRESA_A) Then
'set      strCFOP_ITEM
      If Trim(ST_PRODUTO_A) = 60 Then
         'CFOP 5102 - Venda de mercadoria adquirida ou recebida de terceiros
         'CFOP 5405 - Venda de mercadoria adquirida/recebida de terceiros em operação _
                      com mercadoria sujeita ao regime de substituição tributária, na condição de _
                      contrib substituído
 
'portanto o que vai diferenciar se será um codigo ou outro será a mercadoria em
'si...se ela é substituiçao tributaria ou nao...se for varias mercadorias vc tem que
'verificar uma por uma pra saber.

         strCFOP_ITEM = "5405"
         'Else: strCFOP_ITEM = CFOP_SAIDA_DENTRO_UF_N                     'cfop de venda dentro do estado
      End If

      If RstTemp.State = 1 Then _
         RstTemp.Close

      SQL = "select * from CFOP WITH (NOLOCK)"
      SQL = SQL & " Where CFOP_ID = '" & Trim(strCFOP_ITEM) & "'"
      RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If RstTemp.EOF Then
         If RstTemp.State = 1 Then _
            RstTemp.Close

         If rstProduto.State = 1 Then _
            rstProduto.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP_ITEM & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      'if rstTEMP!Tipo = 0 then 'Dentro do Estado
      VALOR_ICMS_PRODUTO = ((VALOR_TOTAL_ITEM_N * RstTemp!PERC_ICMS) / 100)
      VALOR_PERC_ICMS_N = RstTemp!PERC_ICMS

      If RstTemp.State = 1 Then _
         RstTemp.Close
   End If

   'FORA DO ESTADO
   If UCase(UF_CLIENTE_A) <> UCase(UF_EMPRESA_A) Then
      If Trim(ST_PRODUTO_A) = 60 Then
         strCFOP_ITEM = "6403"  'Fixo por enquanto
         '6403 Venda de mercadoria adquirida ou recebida de terceiros em operação _
               com mercadoria sujeita ao regime de substituição tributária, _
               na condição de contribuinte substituto _
               Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros, _
               na condição de contribuinte substituto, em operação com mercadorias sujeitas _
               ao regime de substituição tributária.

         strCFOP_ITEM = "6404"
         '6404 Venda de mercadoria sujeita ao regime de substituição tributária, _
               cujo imposto já tenha sido retido anteriormente _
               Classificam-se neste código as vendas de mercadorias sujeitas ao regime de substituição tributária, _
               na condição de substituto tributário, exclusivamente nas hipóteses em que o _
               imposto já tenha sido retido anteriormente

         Else: strCFOP_ITEM = CFOP_SAIDA_FORA_UF_N                  'cfop de venda fora do estado do estado
      End If

      SQL = "select * from CFOP WITH (NOLOCK)"
      SQL = SQL & " Where CFOP_ID = '" & Trim(strCFOP_ITEM) & "'"
      RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If RstTemp.EOF Then
         If RstTemp.State = 1 Then _
            RstTemp.Close

         MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP_ITEM & vbCrLf & "Não é possivel continuar a processar"
         'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
         Exit Sub
      End If

      If Trim(Len(CNPJCPF_CLIENTE_A)) > 11 Then ' Se for pessoa juridica
         VALOR_ICMS_PRODUTO = ((VALOR_TOTAL_ITEM_N * RstTemp!PERC_ICMS) / 100)  'CFOP.P_ICMS_VND_F_UF - verificar se existe
         VALOR_PERC_ICMS_N = RstTemp!PERC_ICMS ' CFOP.P_ICMS_VND_F_UF'duas aliquotas para  o mesmo cfop
         Else ' Pessoa fisica
            VALOR_ICMS_PRODUTO = ((VALOR_TOTAL_ITEM_N * RstTemp!ICMS_PJ_F_UF) / 100)
            VALOR_PERC_ICMS_N = RstTemp!ICMS_PJ_F_UF
      End If

      If RstTemp.State = 1 Then _
         RstTemp.Close
   End If

   'HOJE 12/06/2006 22:00
   'FALTA VERIFICAR SE EXISTE DUAS ALIQUOTAS PARA O MESMO CFOP
   'FALTA GRAVAR OS DADOS CORRETAMENTE NA TABELA
   'FALTA VER O LANCE ABAIXO
   
   'Ver depois com o emanoel para que estes campos
   'se for necessarario mesmo, acho que criarei um campo asc de tamanho x
   ' vou appendando os CFOPS que existir separando-os com com um ';"
   'farei uma funcao para tratar os cfops appendando depois
   '   //Testa Cfop para Cabeca!
   '   if PRODUTOS.COD_TRIBUTACAO eq 60 begin
   '      if CIDADE.UF eq DOCUMENT.UF begin
   '         move 5405                               To   CFOP1_D
   '      End
   '      if CIDADE.UF ne DOCUMENT.UF move 6403      To   CFOP1_F
   '   End
   '   if PRODUTOS.COD_TRIBUTACAO ne 60 begin
   '      if CIDADE.UF eq DOCUMENT.UF begin
   '          move CFOP.VND_MERC_D_UF                To   CFOP_D
   '      End
   '      if CIDADE.UF ne DOCUMENT.UF move CFOP.VND_MERC_F_UF;
   '                                                 To   CFOP_F
   '   End

   'If Not isnull(rstProduto!PERCIVA) Then PERC_IVA_N = rstProduto!PERCIVA

   If VALOR_BASE_ICMS_N = 0 Then _
      VALOR_PERC_ICMS_N = 0

   If rstProduto.State = 1 Then _
      rstProduto.Close

'28/03/2017 VERIFICAR SE É ASSIM MESMO:
'QUANDO CLIENTE É CONSUMIDOR FINAL NÃO PASSA NO SEFAZ O PRODUTO COMO SUBSTITUIÇÃO TRIBUTÁRIA
'DAI MUDO AQUI MANUALMENTE A ST DO ITEM PARA 00-TRIBUTADO INTEGRALMENTE
If Trim(UCase(CCE_CLIENTE_A)) = "ISENTO" Or Trim(CCE_CLIENTE_A) = "" Then _
   ST_PRODUTO_A = "00"

'VAI GRAVAR PEDIDO
   If RstTemp.State = 1 Then _
      RstTemp.Close

   SQL = "select PEDIDO.pedido_id,pedidoitem.produto_id from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"

   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where PEDIDO.PEDIDO_ID = " & txtPedido.Text
   SQL = SQL & " And CODG_PRODuto = '" & Trim(txtProduto.Text) & "'"
   SQL = SQL & " and pedidoitem.status <> 'C' "

   RstTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not RstTemp.EOF Then
      PEDIDO_ID_N = RstTemp.Fields(0).Value
      PRODUTO_ID_N = RstTemp.Fields(1).Value

      If RstTemp.State = 1 Then _
         RstTemp.Close

      SQL = "UPDATE PEDIDOITEM SET "
      SQL = SQL & " VlrBaseIcms = " & tpMOEDA(VALOR_BASE_ICMS_N)
      SQL = SQL & ", PERCICMS = " & tpMOEDA(VALOR_PERC_ICMS_N)
      SQL = SQL & ", VlrIcms = " & tpMOEDA(VALOR_ICMS_PRODUTO)
      SQL = SQL & ", VLRBASEICMSSUBST = " & tpMOEDA(VALOR_BASE_ICMS_SUBST_N)
      SQL = SQL & ", PERCICMSSUBST = " & tpMOEDA(VALOR_PERC_ICMS_SUBST_N)
      SQL = SQL & ", VLRICMSSUBST = " & tpMOEDA(VALOR_ICMS_PRODUTO_SUBST_N)
      SQL = SQL & ", cfop_id = '" & Trim(strCFOP_ITEM) & "'"
      SQL = SQL & ", STRIBUTARIA = '" & ST_PRODUTO_A & "'"

      SQL = SQL & " Where pedido_id = " & PEDIDO_ID_N
      SQL = SQL & " and produto_id = " & PRODUTO_ID_N

      CONECTA_RETAGUARDA.Execute SQL
   End If
   If RstTemp.State = 1 Then _
      RstTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREPARA_TRIBUTACAO_PRODUTO"
End Sub

Sub ABRE_PEDIDO()
'On Error GoTo ERRO_TRATA

   If Trim(txtPedido.Text) = "" Then
      txtPedido.Enabled = False

      If Trim(cmbFaturaAUX.Text) = "" Then
         cmbFaturaAUX.Text = 9999
         cmbFatura.Text = "A Vista"
      End If

      If Trim(cmbVendedorAUX.Text) = "" Then
         cmbVendedor.Text = "BALCAO"
         cmbVendedorAUX.Text = 0
      End If

      txtCNPJCPF.PromptInclude = False
      If txtCNPJCPF.Text = "" Then
         txtCNPJCPF.Text = "99999999999"
         If Trim(txtNome.Text) = "" Then _
            txtNome.Text = "Consumidor Final"
      End If

      QUALIFICA_VENDEDOR
      GERA_PEDIDO_ID

      txtPedido.Text = PEDIDO_ID_N
      Else
         If Not IsNumeric(txtPedido.Text) Then
            GERA_PEDIDO_ID

            txtPedido.Text = PEDIDO_ID_N
         End If
   End If

   If Trim(txtPedido.Text) <> "" Then
      If IsNumeric(txtPedido.Text) Then
         txtPedido.Enabled = True
            PEDIDO_ID_N = txtPedido.Text
         txtPedido.Enabled = False

         VALIDA_PEDIDO_ID txtPedido.Text
         SETA_GRID
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ABRE_PEDIDO"
End Sub

Sub VAI_VENDA()
'On Error GoTo ERRO_TRATA

   INDR_GRAVA = False
   If Trim(txtPedido.Text) <> "" Then
      PEDIDO_ID_N = txtPedido.Text
      Else
         MsgBox "Digite Numero da Requisicao para gravar!"
         Exit Sub

         ABRE_PEDIDO
   End If

   txtCNPJCPF.PromptInclude = False
   CHECA_CLIENTE txtCNPJCPF.Text, txtNome.Text
   GERA_VENDA

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select status from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where PEDIDO_ID = " & PEDIDO_ID_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If IsNull(TabTemp.Fields(0).Value) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus

         Exit Sub
      End If
      If TabTemp.Fields(0).Value <= 2 Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         FraSeq.Enabled = True
         txtProduto.Enabled = True
         txtProduto.SetFocus

         Exit Sub
      End If
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
   '===================================

   LIMPA_TUDO
   ABRE_PEDIDO

   FraSeq.Enabled = True
   txtProduto.Enabled = True
   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VAI_VENDA"
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ERRO_TRATA

    ' Quando clicar uma vez
    ' atribui o valor selecionado
    'AtribuiValorCelula
    'OcultarControles

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_Click"
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

   txtProduto.Enabled = True

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MSFlexGrid1_DblClick"
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyF7
         
      Case vbKeyF2      'Editar ao pressionar F2
         ExibirCelula
      Case vbKeyDelete  'Excluir linhas selecionadas
         If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) Then
            If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then
               If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then
                  If Not IsNull(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)) Then
                     If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) <> "" Then                'codg Produto
                        If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)) Then             'pedido_id
                           If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)) Then          'seq_id
                              If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)) Then       'produto_id
                                 txtProduto.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
                                 txtSeq.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
                                 EXCLUIR_ITEM Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Case vbKeyF12
         'frmobs.Show 1
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

      Dim TabDig              As New ADODB.Recordset
      Dim QTDE_RETIDO_ESTORNO As Double

      QTDE_RETIDO_ESTORNO = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)

      AtribuiValorCelula
      'ProximaCelula
      OcultarControles

'==========ATUALIZAR GRID colunas
'3 = qtde
'4 = valor venda
'5 = desconto

      QTDE_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 3)
      VALOR_ITEM_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 4)
      VALOR_DESCONTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 5)
      SEQ_ID_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 11)
      'PRECO_CUSTO_N = 0 & MSFlexGrid1.TextMatrix(LastRow, 8)
      CODG_PRODUTO_A = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 0))
      PRODUTO_ID_N = "" & Trim(MSFlexGrid1.TextMatrix(LastRow, 12))

      If QTDE_N > 0 And VALOR_ITEM_N > 0 And VALOR_DESCONTO_N >= 0 And SEQ_ID_N > 0 Then

         'MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * Qtde_N) - VALOR_DESCONTO_N), strFormatacao2Digitos)  'total item
         MSFlexGrid1.TextMatrix(LastRow, 6) = Format(((VALOR_ITEM_N * QTDE_N)), strFormatacao2Digitos)  'total item
         'lucro MSFlexGrid1.TextMatrix(LastRow, 9) = Format(((VALOR_ITEM_N - PRECO_CUSTO_N) * QTDE_N - VALOR_DESCONTO_N), strFormatacao2Digitos)

         If INDR_ESTQ_NEGATIVO = False Then
            QTDE_ESTOQUE_N = TRAZ_QTDE_ESTOQUE(ESTABELECIMENTO_ID_N, PRODUTO_ID_N)

            If QTDE_ESTOQUE_N < 0 Then
               Beep
               MsgBox "Quantidade pedida maior que quantidade existente no estoque, não permitido.", vbOKOnly, "Atenção."
               txtQtde.SetFocus
               Exit Sub
            End If
         End If

         If TabDig.State = 1 Then _
            TabDig.Close

         SQL = "select PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, PEDIDOITEM.PRODUTO_ID, "
         SQL = SQL & " produto.CODG_PRODuto, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM,"
         SQL = SQL & " PEDIDOITEM.PERC_DESC , PEDIDOITEM.Valor_Desconto, PEDIDOITEM.Status, "
         SQL = SQL & " PEDIDOITEM.PRECO_CUSTO"
         SQL = SQL & " from PEDIDO WITH (NOLOCK) "
         SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK) "
         SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
         SQL = SQL & " INNER JOIN PRODUTO "
         SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
         SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

         SQL = SQL & " where PEDIDO.PEDIDO_ID = " & txtPedido.Text
         SQL = SQL & " and seq_id = " & SEQ_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         SQL = SQL & " and pedidoitem.status <> 'C' "

         TabDig.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

         If Not TabDig.EOF Then
            SQL = "update PEDIDOITEM set "
            SQL = SQL & " QTD_PEDIDA = " & tpMOEDA(QTDE_N)
            SQL = SQL & ",Valor_Item = " & tpMOEDA(VALOR_ITEM_N)
            SQL = SQL & ",Valor_Desconto = " & tpMOEDA(VALOR_DESCONTO_N)
            SQL = SQL & ",peso_item = " & tpMOEDA(QTDE_N)

            SQL = SQL & " where pedido_id = " & TabDig.Fields("pedido_id").Value
            SQL = SQL & " and seq_id = " & SEQ_ID_N
            CONECTA_RETAGUARDA.Execute SQL

            QTDE_RETIDO_ESTORNO = 0

            SETA_GRID
         End If
         If TabDig.State = 1 Then _
            TabDig.Close
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

'&H80C0FF = LARANJA
'&H8000000F = CINZA
'&HFF& = VERMELHO
'vbBlack 0x0
'vbRed 0xFF
'vbGreen 0xFF00
'vbYellow 0xFFFF
'vbBlue 0xFF0000
'vbMagenta 0xFF00FF
'vbCyan 0xFFFF00
'vbWhite 0xFFFFFF

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



Sub QUALIFICA_VENDEDOR()
'On Error GoTo ERRO_TRATA

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbVendedor.Enabled = True
      Else
         If TabUSU.State = 1 Then _
            TabUSU.Close

         SQL = "select logon from USUARIO WITH (NOLOCK)"
         SQL = SQL & " where usuario_id = " & USUARIO_ID_N
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then
            cmbVendedor.Enabled = False

            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close

            CRITERIO_A = Chr$(39) & Trim(TabUSU!Logon) & "%" & Chr(39)
            SQL = "select descricao, vendedor_id from vwVendedor WITH (NOLOCK)"
            SQL = SQL & " where descricao like " & CRITERIO_A
            TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabVENDEDOR.EOF Then
               cmbVendedor.Text = TabVENDEDOR!DESCRICAO
               cmbVendedorAUX.Text = TabVENDEDOR!VENDEDOR_ID
            End If
            If TabVENDEDOR.State = 1 Then _
               TabVENDEDOR.Close
         End If
         If TabUSU.State = 1 Then _
            TabUSU.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "QUALIFICA_VENDEDOR"
End Sub

Sub GERA_IMPRESSAO()
'On Error GoTo ERRO_TRATA

   If txtPedido.Text <> "" Then
      PEDIDO_ID_N = txtPedido.Text
      Else: PEDIDO_ID_N = InputBox(SQL3, "Informe número de Pedido a ser impressa ")
   End If

   FORMULA_REL = "{vwRelVenda.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   FORMULA_REL = FORMULA_REL & " and {vwRelVenda.pedido_id} = " & PEDIDO_ID_N
   FORMULA_REL = FORMULA_REL & " and {vwRelVenda.statusitem} <> 'C' "

   If chkImp.Value = 1 Then _
      ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "rel_pedido_venda.rpt"
   If CNPJ_EMPRESA_N = "15333554000188" Then _
      Nome_Relatorio = "pedido_shf.rpt"

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_IMPRESSAO"
End Sub

Sub VALIDA_PEDIDO_ID(NUMR_PEDIDO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   Dim TabPed  As New ADODB.Recordset
   CRITERIO_A = ""

   If TabPed.State = 1 Then _
      TabPed.Close
'aqui não valida por cpu e establecimento, numero sequencial de pedido
   SQL = "select * from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where pedido_id = " & NUMR_PEDIDO_ID_N
   TabPed.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPed.EOF Then
      bolRequisicaoJaExiste = True
      INDR_VENDA = True

      MOSTRA_DADOS_PEDIDO Trim(TabPed.Fields("CGCCPF").Value), _
                          Trim(TabPed.Fields("VENDEDOR_ID").Value), _
                          Trim(TabPed.Fields("tipovenda_id").Value), _
                          Trim(TabPed.Fields("nome_cliente").Value)

      txtDtEmis.PromptInclude = False
      txtDtEmis.Text = TabPed!dt_req
      txtDtEmis.PromptInclude = True

      If TabPed!Status = 9 Then
         MsgBox "Pedido cancelada, impossível alterar !!!"
         Exit Sub
         Else '1=ORÇAMENTO;2=GERADO;3=EMITIDA COM NOTA;4=EMITIDA COM CUPOM;5=ARECEBER;7=ECF/NF;9=CANCELADO
            If (TabPed!Status = 3 Or TabPed!Status = 5) Then
               If TabPed!Status = 3 Then
                  PERGUNTA "Nota Processada para este pedido.", vbNo, "Venda NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     If TabPed.State = 1 Then _
                        TabPed.Close

                     Else
                        FraSeq.Enabled = False
                        'LIMPA_BODY
                        'LIMPA_TUDO
                   End If
                   Exit Sub
               End If
               If TabPed!Status = 5 Then
                  PERGUNTA "Venda ja Faturada, Deseja imprimir ?", vbYesNo + 32, "Venda NFE", "DEMO.HLP", 1000
                  If RESPOSTA = vbYes Then
                     GERA_IMPRESSAO
                     Else
                        FraSeq.Enabled = False
                        'LIMPA_BODY
                        'LIMPA_TUDO
                   End If
               End If
               Exit Sub
            End If
            If TabPed!Status = 4 Then
               MsgBox "Permitido somente consulta, cupom fiscal emitido."
               Exit Sub
            End If
      End If
   End If
   If TabPed.State = 1 Then _
      TabPed.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "VALIDA_PEDIDO_ID"
End Sub

Private Sub GERA_VENDA()
'On Error GoTo ERRO_TRATA

   PEDIDO_ID_N = txtPedido.Text
   txtCNPJCPF.PromptInclude = False
   CNPJCPF_A = txtCNPJCPF.Text

   If Trim(txtNome.Text) = "" Then _
      txtNome.Text = "Consumidor Final"

   If Trim(cmbVendedorAUX.Text) = "" Then _
      cmbVendedorAUX.Text = 0

   'atualizando desconto na cabeça
   SQL = "UPDATE PEDIDO SET "
   SQL = SQL & " Valor_desconto = 0"
   SQL = SQL & " , Perc_desc = 0"
   SQL = SQL & " , cgccpf = '" & Trim(CNPJCPF_A) & "'"
   SQL = SQL & " , nome_cliente = '" & Trim(txtNome.Text) & "'"
   SQL = SQL & " , status = 2"
   SQL = SQL & " , USUARIO_LIBERA_VENDA = " & USU_LIBERA_VENDA_N
   SQL = SQL & " , vendedor_id = " & cmbVendedorAUX.Text

   SQL = SQL & " where pedido_id = " & txtPedido.Text
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

   If RECEBE_PEDIDO_VENDA = True Then

      FAZ_RECEBIMENTO
      Else
         txtPedido.Text = ""
         ABRE_PEDIDO
   End If

   'rotina que trabalha com a comanda eletronica
   If Trim(lblComanda.Caption) <> "" Then _
      lblComanda.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_VENDA"
End Sub
Private Sub EXCLUIR_ITEM(CODG_PRODUTO_A As String, PEDIDO_ID_N As Long, SEQ_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Trim(PEDIDO_ID_N) > 0 And Trim(SEQ_ID_N) > 0 And Trim(CODG_PRODUTO_A) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select PEDIDOITEM.*, PRODUTO.DESCRICAO, PRODUTO.FAMILIAPRODUTO_ID, "
      SQL = SQL & " PRODUTO.PRECO_VENDA, PRODUTO.PRECO_CUSTO, Produto.Situacao_Tributaria"
      SQL = SQL & " from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
      SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

      SQL = SQL & " where PEDIDOITEM.PEDIDO_ID = " & PEDIDO_ID_N
      SQL = SQL & " and PEDIDOITEM.seq_id = " & SEQ_ID_N
      SQL = SQL & " and codg_produto = '" & Trim(CODG_PRODUTO_A) & "'"
      'SQL = SQL & " and pedidoitem.status <> 'C' "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         Msg = "Deseja cancelar esse item ?  " & Trim(TabTemp.Fields("descricao").Value)
         Style = vbYesNo + 32
         Title = "Atenção."
         Help = "DEMO.HLP"
         Ctxt = 1000
         RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
         If RESPOSTA = vbYes Then
            If TabProduto.State = 1 Then _
               TabProduto.Close

            VALOR_TOTAL_N = Format(VALOR_TOTAL_N - (TabTemp!Valor_Item * TabTemp!QTD_PEDIDA), "##,##0.00")

            'BAIXA_RETIDO  TabTemp!QTD_PEDIDA, TabTemp.Fields("produto_id").Value

            If TabTemp.State = 1 Then _
               TabTemp.Close

            SQL = "Delete from PEDIDOITEM "
            'SQL = "update PEDIDOITEM set "
            'SQL = SQL & " status = 'C' "
            SQL = SQL & " Where pedido_id = " & PEDIDO_ID_N
            SQL = SQL & " and seq_id = " & SEQ_ID_N
            'SQL = SQL & " and tipo_reg = 'PC' "
            CONECTA_RETAGUARDA.Execute SQL

            If TabTemp.State = 1 Then _
               TabTemp.Close

            LIMPA_BODY
            lblTotal.Caption = "Total Orçamento = " & Format(VALOR_TOTAL_N, "currency")
   
            GRAVA_CABECA "R"
            SETA_GRID
            Else
               If TabTemp.State = 1 Then _
                  TabTemp.Close
         End If
         Else: MsgBox "Produto não encontrado."
      End If
      Else: MsgBox "Informe código produto."
   End If
      FraSeq.Enabled = True
      txtProduto.Enabled = True

   txtProduto.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "EXCLUIR_ITEM"
End Sub


Private Sub MOSTRA_DADOS_PEDIDO(CNPJCPF_PEDIDO_A As String, _
                                VENDEDOR_PEDIDO_ID_N As Long, _
                                TIPO_VENDA_PEDIDO_ID_N As Long, _
                                NOME_CLI_PEDIDO_A As String)
'On Error GoTo ERRO_TRATA

   Dim TabCons  As New ADODB.Recordset

   CNPJCPF_PEDIDO_A = Trim(CNPJCPF_PEDIDO_A)
   txtCNPJCPF.PromptInclude = False
   txtCNPJCPF.Text = CNPJCPF_PEDIDO_A

   If MULT_EMPRESA_B = False Then
      'MOSTRA VENDEDOR
      If Not IsNull(VENDEDOR_PEDIDO_ID_N) Then
         cmbVendedorAUX.Text = VENDEDOR_PEDIDO_ID_N

         If TabCons.State = 1 Then _
            TabCons.Close
         SQL = "select descricao from vwVendedor WITH (NOLOCK)"
         SQL = SQL & " where vendedor_id = " & cmbVendedorAUX.Text
         TabCons.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCons.EOF Then _
            cmbVendedor.Text = "" & Trim(TabCons.Fields("descricao").Value)
         If TabCons.State = 1 Then _
            TabCons.Close
      End If
   End If

   'MOSTRA CLIENTE
   If TabCons.State = 1 Then _
      TabCons.Close

   SQL = "select nome,status from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " where cgccpf = '" & CNPJCPF_PEDIDO_A & "'"
   SQL = SQL & " and status = 'A'"
   TabCons.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCons.EOF Then
      If CNPJCPF_PEDIDO_A = "99999999999" Then
         If Not IsNull(NOME_CLI_PEDIDO_A) Then
            If Trim(txtNome.Text) = "" Then _
               txtNome.Text = NOME_CLI_PEDIDO_A
            Else
               If Trim(txtNome.Text) = "" Then _
                  txtNome.Text = TabCons!NOME
         End If
         Else
            If Trim(txtNome.Text) = "" Then _
               txtNome.Text = TabCons!NOME
      End If
   End If
   If TabCons.State = 1 Then _
      TabCons.Close

   SQL = "select nome_cliente from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " where PEDIDO_Id = " & PEDIDO_ID_N
   TabCons.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCons.EOF Then _
      If Not IsNull(TabCons.Fields(0).Value) Then _
         txtNome.Text = Trim(TabCons.Fields(0).Value)
   If TabCons.State = 1 Then _
      TabCons.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_DADOS_PEDIDO"
End Sub

Private Sub FAZ_RECEBIMENTO()
'On Error GoTo ERRO_TRATA

   If PEDIDO_ID_N <= 0 Then
      MsgBox "Pedido inválido, verifique !!!"
      Exit Sub
   End If

   Dim TabPedido     As New ADODB.Recordset
   Dim INDR_PERGUNTA As Boolean

   If PEDIDO_ID_N > 0 Then
      INDR_RECEITA = 1

      If INDR_FORM_ABERTO = True Then
         Unload frmFatura
         INDR_FORM_ABERTO = False
      End If
'===================================
      If Trim(cmbFaturaAUX.Text) = "" Then
         cmbFaturaAUX.Text = 9999
         cmbFatura.Text = "A Vista"
      End If

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select contabiliza from TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " where tipovenda_id = " & cmbFaturaAUX.Text
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields("contabiliza").Value) Then
            If TabTemp.Fields("contabiliza").Value = True Then
               If TabTemp.State = 1 Then _
                  TabTemp.Close

frmFatura.Show 1

               Else
                  SQL = "update PEDIDO set "
                  SQL = SQL & "status = 6 " 'não contabiliza
                  SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                  SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
                  CONECTA_RETAGUARDA.Execute SQL

'essa rotina tem que testar quando a condição da empresa for esta, não contabiliza
                  If Trim(lblComanda.Caption) <> "" Then _
                     lblComanda.Visible = True
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
'===================================
      If INDR_CONTROLA_ESTOQUE = False Then _
         Exit Sub

      If TabPedido.State = 1 Then _
         TabPedido.Close

      SQL = "select * from PEDIDO WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPedido.EOF Then
         PEDIDO_ID_N = TabPedido.Fields("pedido_id").Value
         If TabPedido!Status = 5 Then
            CNPJCPF_A = Trim(TabPedido!CGCCPF)
'=============================================================================
'comanda
            SQL = "delete pedidoTEMP where pedido_id = " & txtPedido.Text
            CONECTA_RETAGUARDA.Execute SQL

'=============NFC-E
            If USA_DOC_FISCAL = True Then
               If USA_NFC_E = True Then  'se usa NFC-e ENTRA AQUI
                  RESPOSTA = ""
                  Msg = ""
                  If INDR_VENDA_CARTAO = True Then
                     RESPOSTA = vbYes
                     Else
                        Msg = "Deseja Gerar Cupom Eletrônico ?"
                        PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                  End If

                  If RESPOSTA = vbYes Then
                     frmDISPLAYEMISSOR.ROTINA_NFC
                  End If
               End If

               txtCNPJCPF.PromptInclude = False
               If Trim(txtCNPJCPF.Text) <> "99999999999" Then
                  If USA_NFe = True And INDR_CAIXA = False Then
                     Msg = "Deseja Gerar Nota Fiscal Eletrônica ?"
                     PERGUNTA Msg, vbYesNo + 32, "Faturamento", "DEMO.HLP", 1000
                     If RESPOSTA = vbYes Then
                        If TabPedido.Fields("Status").Value = 5 Or TabPedido.Fields("Status").Value = 7 Then
                           CRITERIO_A = PEDIDO_ID_N
                           TIPO_NFe_GERAR = "R"
                           frmNOTAGERA.Show 1
                        End If
                     End If
                  End If
               End If
            End If   'If USA_DOC_FISCAL = True Then
         End If
      End If
      If TabPedido.State = 1 Then _
         TabPedido.Close

'===================================
      If INDR_CONTROLA_ESTOQUE = True Then
         '====================
            ATUALIZA_ESTOQUE 0, PEDIDO_ID_N
         '====================
      End If
'===================================
   End If
   If TabPedido.State = 1 Then _
      TabPedido.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FAZ_RECEBIMENTO"
End Sub
