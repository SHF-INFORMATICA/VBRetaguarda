VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmATUALIZACAO 
   Caption         =   "Atualizações"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ATUALIZACAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   10560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "implantação dochefe"
      Height          =   375
      Left            =   8040
      TabIndex        =   32
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdUF 
      Caption         =   "Tabela de códigos das Unidades Federativas/Estados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   31
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command30 
      Caption         =   "INDUSTRIA"
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
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Acerto Desconto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   29
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      Caption         =   "IMPORTA JHOU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   28
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      Caption         =   "CRIA CARTÃO COMANDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   27
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command23 
      Caption         =   "TABELA ERRO"
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
      Left            =   10680
      TabIndex        =   26
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command17 
      Caption         =   "FORNECEDOR X PESSOA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   25
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton CMDSU 
      Caption         =   "impora clientes RESTAURANTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   24
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excluir nf Global"
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
      Left            =   10680
      TabIndex        =   23
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton Command27 
      Caption         =   "IMPRESSORA FISCAL"
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
      Left            =   5400
      TabIndex        =   22
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Command12 
      Caption         =   "LEI DO IMPOSTO 12.741"
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
      Left            =   2760
      TabIndex        =   21
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ACERTO PESO/QTDE ITENS"
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
      Left            =   10680
      TabIndex        =   20
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Pedidos Cancelados Títulos Cancelados"
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
      Left            =   10680
      TabIndex        =   19
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      Caption         =   "TABELA CHAMADO"
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
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ZERA BANCO GLOBAL"
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
      Left            =   10680
      TabIndex        =   17
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "AT NOME CLIENTE PEDIDO"
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
      Left            =   10680
      TabIndex        =   16
      Top             =   4800
      Width           =   3135
   End
   Begin VB.CommandButton cmdImpRel 
      Caption         =   "PROGRAMA/IMPRESSORA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ESTOQUE CHECA"
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
      Left            =   10680
      TabIndex        =   14
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton Command26 
      Caption         =   "TROCA PRODUTO"
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
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command25 
      Caption         =   "BANCO,AGENCIA,CONTA,CHEQUE"
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
      Left            =   5400
      TabIndex        =   12
      Top             =   1200
      Width           =   2415
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   -120
      TabIndex        =   11
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   661
      _Version        =   262144
      Caption         =   "CRIAR BANCOS (MEGASIM,GLOBAL)"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.CommandButton cmdCCE 
      Caption         =   "Inscrição Estadual Municipal"
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
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdCOMISSAO 
      Caption         =   "TABELA COMISSÃO"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Lote Produto"
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
      Left            =   8040
      TabIndex        =   8
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command21 
      Caption         =   "BANCO SUSPEITO"
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
      Left            =   10680
      TabIndex        =   7
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton Command20 
      Caption         =   "ZERA BANCO SHFSYS"
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
      Left            =   10680
      TabIndex        =   6
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command18 
      Caption         =   "TABELA PROGRAMA"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton Command16 
      Caption         =   "ATUALIZA CEP"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdFornecedor 
      Caption         =   "Tabela Fornecedor"
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
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdCupom 
      Caption         =   "Atualiza Tabela CUPOM"
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
      Left            =   5400
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Titulos data baixa status aberto"
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
      Left            =   10680
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton cmdKardex 
      Caption         =   "Atualização Kardex"
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
      Left            =   5400
      TabIndex        =   0
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   12
      X1              =   0
      X2              =   13800
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   4
      X1              =   10560
      X2              =   10560
      Y1              =   8400
      Y2              =   -360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   3
      X1              =   7920
      X2              =   7920
      Y1              =   8400
      Y2              =   -360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Index           =   2
      X1              =   5280
      X2              =   5280
      Y1              =   8400
      Y2              =   -360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   9240
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Index           =   0
      X1              =   2640
      X2              =   2640
      Y1              =   8400
      Y2              =   -360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   11
      X1              =   0
      X2              =   13800
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   10
      X1              =   0
      X2              =   13800
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   9
      X1              =   0
      X2              =   13800
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   8
      X1              =   0
      X2              =   13800
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   7
      X1              =   0
      X2              =   13800
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   6
      X1              =   0
      X2              =   13800
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   5
      X1              =   0
      X2              =   13800
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   4
      X1              =   0
      X2              =   13800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   3
      X1              =   0
      X2              =   13800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   13800
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   13800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   13800
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmATUALIZACAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim CNPJCPF    As String
   Dim DESCRICAO  As String
   Dim xl         As New Excel.Application
   Dim xlw        As Excel.Workbook

Private Sub cmdUF_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "UF", "U") = False Then
      SQL = "CREATE TABLE [dbo].[UF]("
      SQL = SQL & " [cUF] [bigint] NOT NULL,"
      SQL = SQL & " [DESCRICAO] [NVARCHAR](100) NOT NULL,"
      SQL = SQL & " [ESTADO] [NVARCHAR](100) NOT NULL"
      SQL = SQL & " CONSTRAINT [PK_UF] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([cUF] ASC)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
      SQL = SQL & " ON [PRIMARY] ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 12
         SQL = SQL & ",'Acre'"
         SQL = SQL & ",'AC'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 27
         SQL = SQL & ",'Alagoas'"
         SQL = SQL & ",'AL'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 16
         SQL = SQL & ",'Amapá'"
         SQL = SQL & ",'AP'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   
      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 13
         SQL = SQL & ",'Amazonas'"
         SQL = SQL & ",'AM'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 29
         SQL = SQL & ",'Bahia'"
         SQL = SQL & ",'BA'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 23
         SQL = SQL & ",'Ceará'"
         SQL = SQL & ",'CE'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 53
         SQL = SQL & ",'Distrito Federal'"
         SQL = SQL & ",'DF'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 32
         SQL = SQL & ",'Espírito Santo'"
         SQL = SQL & ",'ES'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 52
         SQL = SQL & ",'Goiás'"
         SQL = SQL & ",'GO'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 21
         SQL = SQL & ",'Maranhão'"
         SQL = SQL & ",'MA'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 51
         SQL = SQL & ",'Mato Grosso'"
         SQL = SQL & ",'MT'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 50
         SQL = SQL & ",'Mato Grosso do Sul'"
         SQL = SQL & ",'MS'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 31
         SQL = SQL & ",'Minas Gerais'"
         SQL = SQL & ",'MG'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 15
         SQL = SQL & ",'Pará'"
         SQL = SQL & ",'PA'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 25
         SQL = SQL & ",'Paraíba'"
         SQL = SQL & ",'PB'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 41
         SQL = SQL & ",'Paraná'"
         SQL = SQL & ",'PR'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 26
         SQL = SQL & ",'Pernambuco'"
         SQL = SQL & ",'PE'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 22
         SQL = SQL & ",'Piauí'"
         SQL = SQL & ",'PI'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 33
         SQL = SQL & ",'Rio de Janeiro'"
         SQL = SQL & ",'RJ'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 24
         SQL = SQL & ",'Rio Grande do Norte'"
         SQL = SQL & ",'RN'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 43
         SQL = SQL & ",'Rio Grande do Sul'"
         SQL = SQL & ",'RS'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 11
         SQL = SQL & ",'Rondônia'"
         SQL = SQL & ",'RO'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 14
         SQL = SQL & ",'Roraima'"
         SQL = SQL & ",'RR'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 42
         SQL = SQL & ",'Santa Catarina'"
         SQL = SQL & ",'SC'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 35
         SQL = SQL & ",'São Paulo'"
         SQL = SQL & ",'SP'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 28
         SQL = SQL & ",'Sergipe'"
         SQL = SQL & ",'SE'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into UF "
         SQL = SQL & " (cUF,DESCRICAO,ESTADO) "
      SQL = SQL & " values("
         SQL = SQL & 17
         SQL = SQL & ",'Tocantins'"
         SQL = SQL & ",'TO'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "OK"
End Sub

Private Sub Command19_Click()
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select pedido.pedido_id from pedido "
   SQL = SQL & " where tipovenda_id = 1"
   SQL = SQL & " order by pedido_id desc"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
'convenio
      If TabTemp.Fields("tipovenda_id").Value = 1 Then

         valor_tot_venda_n = 0
         'If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select sum(qtd_pedida*valor_item) from PEDIDOITEM "
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then _
            If Not IsNull(TabConsulta.Fields(0).Value) Then _
               valor_tot_venda_n = 0 & TabConsulta.Fields(0).Value
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         'SQL = "update pedido set "
         'SQL = SQL & " perc_desc = 0 "
         'SQL = SQL & " ,VALOR_DESCONTO = 0"
         'SQL = SQL & " ,valor_total = " & tpMOEDA(valor_tot_venda_n)
            'SQL = SQL & " ,valor_recebido = " & tpMOEDA(valor_tot_venda_n)
         'SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         'CONECTA_RETAGUARDA.Execute SQL

'==============================

         SQL = "select (PEDIDOITEM.VALOR_ITEM*PEDIDOITEM.QTD_PEDIDA) as totitem "
         SQL = SQL & " from PEDIDOITEM "
         SQL = SQL & " INNER JOIN PRODUTO "
         SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
         SQL = SQL & " AND PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         SQL = SQL & " and Produto.PERMITE_DESCONTO = 1"
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            If Not IsNull(TabConsulta.Fields(0).Value) Then

               'peguei o desconto somente do item que permite desconto
               VALOR_DESCONTO_N = TabConsulta.Fields(0).Value * 10 / 100
               'valor_tot_venda_n = TabConsulta.Fields(0).Value

               SQL = "update pedido set "

               SQL = SQL & " perc_desc = 10 "
               SQL = SQL & " ,VALOR_DESCONTO = " & tpMOEDA(VALOR_DESCONTO_N)
               SQL = SQL & " ,valor_total = " & tpMOEDA(valor_tot_venda_n)
               SQL = SQL & " ,valor_recebido = " & tpMOEDA(valor_tot_venda_n - VALOR_DESCONTO_N)
'atualizar com o total da venda - o desconto
               SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
               CONECTA_RETAGUARDA.Execute SQL

               SQL = "update itemlancamento set valor_item = " & tpMOEDA(valor_tot_venda_n - VALOR_DESCONTO_N)
               SQL = SQL & " where numr_doc = " & TabTemp.Fields("pedido_id").Value
               SQL = SQL & " and seq = 1"
               CONECTA_RETAGUARDA.Execute SQL

               CMDSU.Caption = TabTemp.Fields("pedido_id").Value

            End If
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close
      End If

      VALOR_DESCONTO_N = 0
      Command19.Caption = TabTemp.Fields("pedido_id").Value
      DoEvents
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

MsgBox "ok"
End Sub

Private Sub Command28_Click()

   CONT_N = 0

   While CONT_N < 100000
      DoEvents
      CONT_N = CONT_N + 1
      Command28.Caption = CONT_N

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from CARTAOBARRA "
      SQL = SQL & " where CARTAOBARRA_ID = " & CONT_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "insert into CARTAOBARRA "
         SQL = SQL & " values("
            SQL = SQL & CONT_N                              'CARTAOBARRA_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N          'ESTABELECIMENTO_ID
            SQL = SQL & "," & CONT_N                        'CODIGO_BARRA
            SQL = SQL & ",'" & "Comanda" & CONT_N & "'"     'DESCRICAO
            SQL = SQL & ",'" & Now & "'"                    '
            SQL = SQL & ",'A'"
         SQL = SQL & " )"
         
Command28.Caption = CONT_N
DoEvents

         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close
   Wend

MsgBox "ok"
End Sub

Private Sub Command29_Click()

   Proc_n = 0
   Novos_n = 0
   At_n = 0

   Set oConn = New ADODB.Connection
   oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                      "FIL=excel 8.0;" & _
                      "DefaultDir=c:\MEGASIM\txt\;" & _
                      "MaxBufferSize=2048;" & _
                      "PageTimeout=5;" & _
                      "DBQ=c:\MEGASIM\txt\imp.xls;"

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   'aabre o recordset pelo nome da planilha
   TabConsulta.Open "[tab$]", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

   TabConsulta.MoveFirst

   If TabConsulta.EOF Then
      MsgBox "Planilha incorreta !!!"
      Exit Sub
   End If

   While Not TabConsulta.EOF
      Proc_n = Proc_n + 1
      CMDSU.Caption = "Processados = " & Proc_n

      If Not IsNull(TabConsulta(0).Value) Then
         If IsNumeric(TabConsulta(0).Value) Then

   'FAMILIA
            NUMR_ID_N = 0
            If Not IsNull(TabConsulta(2).Value) Then
               
                  If TabProduto.State = 1 Then _
                     TabProduto.Close

                  SQL = "select familiaproduto_id from FAMILIAPRODUTO WITH (NOLOCK)"
                  SQL = SQL & " where descricao = '" & Trim(TabConsulta(2).Value) & "'"
                  TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabProduto.EOF Then
                     NUMR_ID_N = MAX_ID("familiaproduto_id", "familiaproduto", "", "", "", "")

                     SQL = "insert into FAMILIAPRODUTO "
                        SQL = SQL & "(FAMILIAPRODUTO_ID,CODG_FAMILIA,DESCRICAO,UNIDADE_MEDIDA,DESC_UNIDADE_MEDIDA,PRODUCAO)"
                     SQL = SQL & " values("
                        SQL = SQL & NUMR_ID_N                                 'FAMILIAPRODUTO_ID
                        SQL = SQL & ",'" & NUMR_ID_N & "'"                    'CODG_FAMILIA
                        SQL = SQL & ",'" & Trim(TabConsulta(2).Value) & "'"   'DESCRICAO
                        SQL = SQL & ",'" & Trim(TabConsulta(4).Value) & "'"   'UNIDADE_MEDIDA
                        SQL = SQL & ",'" & Trim(TabConsulta(4).Value) & "'"   'DESC_UNIDADE_MEDIDA
                        SQL = SQL & ",0"                                      'PRODUCAO
                     SQL = SQL & " )"
                     CONECTA_RETAGUARDA.Execute SQL
                     Else: NUMR_ID_N = TabProduto.Fields(0).Value
                  End If
               
            End If
            If TabProduto.State = 1 Then _
               TabProduto.Close

            SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
            SQL = SQL & " where codg_produto = '" & Trim(TabConsulta(0).Value) & "'"
            TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabProduto.EOF Then
               PRODUTO_ID_N = MAX_ID("produto_ID", "produto", "", "", "", "")
   
               SQL = "insert into PRODUTO "
                  SQL = SQL & "(produto_id,codg_produto,descricao,familiaproduto_id,"
                  SQL = SQL & " unidade_medida,situacao,tipo_prod,preco_custo_anterior,"
                  SQL = SQL & " preco_custo,preco_atacado,preco_venda,dt_cadastro,"
                  SQL = SQL & " preco_varejo_anterior,preco_atacado_anterior,empresa_id,situacao_tributaria,aliquota_icms)"
               SQL = SQL & " values("
                  SQL = SQL & PRODUTO_ID_N                              'produto_id
                  SQL = SQL & ",'" & Trim(TabConsulta(0).Value) & "'"   'codg_produto
                  SQL = SQL & ",'" & Trim(TabConsulta(3).Value) & "'"   'descricao
                  SQL = SQL & "," & NUMR_ID_N                           'familiaproduto_id
                  SQL = SQL & ",'" & Trim(TabConsulta(4).Value) & "'"   'Unidade_Medida
                  SQL = SQL & ",'A' "                                   'SITUACAO
                  SQL = SQL & ",1"                                      'Tipo_Prod
                  SQL = SQL & "," & tpMOEDA(TabConsulta(16).Value)       'PRECO_CUSTO_ANTERIOR,
                  SQL = SQL & "," & tpMOEDA(TabConsulta(16).Value)       'preco_custo
                  SQL = SQL & "," & tpMOEDA(TabConsulta(16).Value)       'preco_atacado
                  SQL = SQL & "," & tpMOEDA(TabConsulta(16).Value)       'preco_venda
                  SQL = SQL & ",'" & Now & "'"                           'dt_cadastro
                  SQL = SQL & "," & tpMOEDA(TabConsulta(16).Value)       'preco_varejo_anterior
                  SQL = SQL & "," & tpMOEDA(TabConsulta(16).Value)       'preco_atacado_anterior
                  SQL = SQL & ",1"
                  SQL = SQL & ",'00'"
                  SQL = SQL & ",17"
               SQL = SQL & " )"
               CONECTA_RETAGUARDA.Execute SQL
               Else
                  SQL = "update produto set "
                  SQL = SQL & " descricao = '" & Trim(TabConsulta.Fields(3).Value) & "'"
                  SQL = SQL & " ,situacao_tributaria='00'"
                  SQL = SQL & " ,aliquota_icms=17 "
                  SQL = SQL & " where codg_produto = '" & Trim(TabConsulta(0).Value) & "'"
                  CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabProduto.State = 1 Then _
               TabProduto.Close
         End If
      End If

      DoEvents
On Error Resume Next
      TabConsulta.MoveNext
Err.Clear
   Wend

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

MsgBox "OK"
End Sub

Private Sub Command30_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "MP", "U") = False Then
      SQL = "CREATE TABLE [dbo].[MP]("
      SQL = SQL & " [PRODUTO_ID_PA] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID_MP] [bigint] NOT NULL,"
      SQL = SQL & " [VALOR] [float] NOT NULL,"
      SQL = SQL & " [QTDE] [float] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_MP] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ("
      SQL = SQL & " [PRODUTO_ID_PA] ASC,"
      SQL = SQL & " [PRODUTO_ID_MP] Asc"
      SQL = SQL & " )"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
      SQL = SQL & " ON [PRIMARY] ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[MP]  WITH CHECK ADD  CONSTRAINT [FK_MP_PRODUTO] FOREIGN KEY([PRODUTO_ID_PA])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[MP] CHECK CONSTRAINT [FK_MP_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[MP]  WITH CHECK ADD  CONSTRAINT [FK_MP_PRODUTO1] FOREIGN KEY([PRODUTO_ID_MP])"
      SQL = SQL & " References [dbo].[Produto]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[MP] CHECK CONSTRAINT [FK_MP_PRODUTO1]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ORDEMPRODUCAO", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ORDEMPRODUCAO]("
      SQL = SQL & " [ORDEMPRODUCAO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
      SQL = SQL & " [DT_BAIXA] [datetime] NOT NULL,"
      SQL = SQL & " [SITUACAO] [nchar](1) NOT NULL,"
      SQL = SQL & " [RESP_ID_CAD] [int] NOT NULL,"
      SQL = SQL & " [RESP_ID_BAIXA] [int] NULL,"
      SQL = SQL & " CONSTRAINT [PK_ORDEMPRODUCAO] PRIMARY KEY CLUSTERED([ORDEMPRODUCAO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAO]  WITH CHECK ADD  CONSTRAINT [FK_ORDEMPRODUCAO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAO] CHECK CONSTRAINT [FK_ORDEMPRODUCAO_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ORDEMPRODUCAOITEM", "U") = False Then
      SQL = "CREATE TABLE [dbo].[ORDEMPRODUCAOITEM]("
      SQL = SQL & " [ORDEMPRODUCAO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID_PA] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID_MP] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE_MP] [float] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_ORDEMPRODUCAOITEM] PRIMARY KEY CLUSTERED("
      SQL = SQL & " [ORDEMPRODUCAO_ID] ASC,[PRODUTO_ID_PA] ASC,[PRODUTO_ID_MP] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_ORDEMPRODUCAOITEM_ORDEMPRODUCAO] FOREIGN KEY([ORDEMPRODUCAO_ID])"
      SQL = SQL & " References [dbo].[ORDEMPRODUCAO]([ORDEMPRODUCAO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAOITEM] CHECK CONSTRAINT [FK_ORDEMPRODUCAOITEM_ORDEMPRODUCAO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_ORDEMPRODUCAOITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID_PA])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAOITEM] CHECK CONSTRAINT [FK_ORDEMPRODUCAOITEM_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_ORDEMPRODUCAOITEM_PRODUTO1] FOREIGN KEY([PRODUTO_ID_MP])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ORDEMPRODUCAOITEM] CHECK CONSTRAINT [FK_ORDEMPRODUCAOITEM_PRODUTO1]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "ok"
End Sub

Private Sub Form_Load()
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
   End Select
End Sub

Private Sub Command17_Click()
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwFornecedor "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      SQL = "update PESSOA set "
         SQL = SQL & " descricao = '" & Trim(TabTemp.Fields("nome").Value) & "'"
         SQL = SQL & " , razao = '" & Trim(TabTemp.Fields("razao_social").Value) & "'"
      SQL = SQL & " where cnpjcpf = '" & TabTemp.Fields("cgccpf").Value & "'"
      CONECTA_RETAGUARDA.Execute SQL

      Command17.Caption = Trim(TabTemp.Fields("razao_social").Value)
      DoEvents

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

MsgBox "ok"
End Sub

Private Sub Command23_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ERRO", "U") = True Then
   
   End If
End Sub

Private Sub Command1_Click()
   Dim VACA_VEIA

   VACA_VEIA = InputBox("Informe Nome Banco de Dados a ser criado.", "SHF INFORMÁTICA")
   If Trim(VACA_VEIA) <> "" Then
      If IsNumeric(VACA_VEIA) Then
         ABRE_BANCO_GLOBAL

         If CONECTA_GLOBAL.State <> 1 Then
            MsgBox "Banco GLOBAL não conectado."
            Exit Sub
         End If

         CONECTA_GLOBAL.Execute "DELETE from MFi010 where mfidoc = '" & VACA_VEIA & "'"

         SQL = "DELETE from MFA010 "
         SQL = SQL & " where mfadoc = '" & VACA_VEIA & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

         CONECTA_GLOBAL.Execute SQL

         If CONECTA_GLOBAL.State = 1 Then _
            CONECTA_GLOBAL.Close

         MsgBox "ok"
      End If
   End If
End Sub

Private Sub cmdCCE_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "IE", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENDERECO_ID", "IE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IE ADD ENDERECO_ID BIGINT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_IE_ENDERECO", "") = False Then
         SQL = "ALTER TABLE [dbo].[IE] WITH CHECK ADD CONSTRAINT [FK_IE_ENDERECO] FOREIGN KEY([ENDERECO_ID])"
         SQL = SQL & " References [dbo].[ENDERECO]([ENDERECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[IE] CHECK CONSTRAINT [FK_IE_ENDERECO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "IM", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENDERECO_ID", "IM") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IM ADD ENDERECO_ID BIGINT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_IM_ENDERECO", "") = False Then
         SQL = "ALTER TABLE [dbo].[IM] WITH CHECK ADD CONSTRAINT [FK_IM_ENDERECO] FOREIGN KEY([ENDERECO_ID])"
         SQL = SQL & " References [dbo].[ENDERECO]([ENDERECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[IM] CHECK CONSTRAINT [FK_IM_ENDERECO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

MsgBox "OK"
End Sub

Private Sub cmdFornecedor_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FORNECEDOR", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "FORNECEDOR") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR ADD PESSOA_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT not null", "FORNECEDOR", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORNECEDOR_ID", "FORNECEDOR") = True Then _
         Alteração_Definição_Campo_Tabela "FORNECEDOR_ID", "BIGINT not null", "FORNECEDOR", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_FORNECEDOR", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR ADD CONSTRAINT pk_FORNECEDOR PRIMARY KEY (FORNECEDOR_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_FORNECEDOR_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[FORNECEDOR]  WITH CHECK ADD  CONSTRAINT [FK_FORNECEDOR_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[FORNECEDOR] CHECK CONSTRAINT [FK_FORNECEDOR_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDO DROP COLUMN EMPRESA_ID      "

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_FORNECEDOR_EMPRESA", "") = True Then
         SQL = "alter table FORNECEDOR "
         SQL = SQL & " drop CONSTRAINT FK_FORNECEDOR_EMPRESA"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "FORNECEDOR") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR DROP COLUMN EMPRESA_ID"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "FORNECEDOR") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORNECEDOR ADD ESTABELECIMENTO_ID INT"
         SQL = "update FORNECEDOR set estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_FORNECEDOR_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[FORNECEDOR]  WITH CHECK ADD  CONSTRAINT [FK_FORNECEDOR_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[FORNECEDOR] CHECK CONSTRAINT [FK_FORNECEDOR_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   MsgBox "Ok, TABELEA FORNECEDOR , ATENÇÃO CRIAR INDICE CNPJCPF  =  " & CONT_N
End Sub

Private Sub cmdCupom_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CUPOM", "U") = True Then
MsgBox "organizar tabela"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CUPOM_ID", "CUPOM") = True Then _
         Alteração_Definição_Campo_Tabela "CUPOM_ID", "BIGINT NOT NULL", "CUPOM", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_CUPOM", "CUPOM") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_CUPOM", "BIGINT NOT NULL", "CUPOM", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "CUPOM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CUPOM ADD PEDIDO_ID BIGINT NOT NULL"
         Else
            If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "CUPOM") = True Then _
               Alteração_Definição_Campo_Tabela "PEDIDO_ID", "BIGINT NOT NULL", "CUPOM", "RETAGUARDA"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CUPOM", "") = False Then
         SQL = "ALTER TABLE CUPOM ADD CONSTRAINT pk_CUPOM PRIMARY KEY (CUPOM_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CUPOM_PEDIDO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CUPOM] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_CUPOM_PEDIDO] "
         SQL = SQL & " FOREIGN KEY([PEDIDO_ID])"
         SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CUPOM] CHECK CONSTRAINT [FK_CUPOM_PEDIDO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CUPOM_IMPRESSORA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CUPOM]  WITH CHECK ADD  CONSTRAINT [FK_CUPOM_IMPRESSORA] FOREIGN KEY([IMPRESSORA_ID])"
         SQL = SQL & " References [dbo].[IMPRESSORA]([IMPRESSORA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CUPOM] CHECK CONSTRAINT [FK_CUPOM_IMPRESSORA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub Command10_Click()
   CRITERIO_A = Command10.Caption

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select pedido_id from PEDIDO "
   SQL = SQL & " where status = 9"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      SQL = "update ITEMLANCAMENTO set "
         SQL = SQL & " Status = 'C'"
         SQL = SQL & ", DT_baixa = null"
         SQL = SQL & ", DT_cancela = " & Now
      SQL = SQL & " where numr_doc = " & TabConsulta.Fields(0).Value
      SQL = SQL & " and status <> 'C'"
      CONECTA_RETAGUARDA.Execute SQL
DoEvents
Command10.Caption = TabConsulta.Fields(0).Value
      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   MsgBox "ok"
   Command10.Caption = CRITERIO_A
   CRITERIO_A = ""
End Sub

Private Sub Command11_Click()

   Dim VALOR_KILO_N As Double

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select valor_item,produto_id,pedido_id,seq_id from PEDIDOITEM "
   SQL = SQL & " order by pedido_id"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      If Not IsNull(TabTemp.Fields(0).Value) Then
         VALOR_KILO_N = 0

         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "select preco_venda from PRODUTO "
         SQL = SQL & " where produto_id = " & TabTemp.Fields(1).Value
         TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabProduto.EOF Then _
            If Not IsNull(TabProduto.Fields(0).Value) Then _
               VALOR_KILO_N = 0 & TabProduto.Fields(0).Value

         If TabProduto.State = 1 Then _
            TabProduto.Close

         SQL = "update PEDIDOITEM set qtd_pedida = " & tpMOEDA(CONVERTE_VALOR_GRAMA(TabTemp.Fields(0).Value, VALOR_KILO_N, 0))
         SQL = SQL & " where pedido_id =  " & TabTemp.Fields(2).Value
         SQL = SQL & " and seq_id = " & TabTemp.Fields(3).Value
         CONECTA_RETAGUARDA.Execute SQL
         DoEvents
         Command11.Caption = TabTemp.Fields(2).Value
      End If

      TabTemp.MoveNext
   Wend
Command11.Caption = "ACERTO PESO/QTDE ITENS"
   If TabTemp.State = 1 Then _
      TabTemp.Close
MsgBox "OK"
End Sub

Private Sub Command12_Click()

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TABNCM", "U") = False Then
      SQL = "CREATE TABLE [dbo].[TABNCM]("
      SQL = SQL & " [CODG_NCM] [nvarchar](8) NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](max) NOT NULL,"
      SQL = SQL & " [ALIQUOTA_NCM] [float] NOT NULL,"
      SQL = SQL & "  CONSTRAINT [PK_TABNCM] PRIMARY KEY CLUSTERED([CODG_NCM] Asc"
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If FSO.FileExists(App.Path & "\TXT\NCMimporta.XLS") Then _
      RODA_IMPORTA_NCM
MsgBox "ok, fim "
End Sub

Private Sub Command16_Click()
   If TabCEP.State = 1 Then _
      TabCEP.Close

   SQL = "select * from CEP "
   SQL = SQL & " order by CEP_id,CIDADE,UF,IBGE_ID "
   TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   While Not TabCEP.EOF

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      SQL = "select * from CEP_vaca "
      SQL = SQL & " where cep_ID = '" & TabCEP.Fields("cep_id").Value & "'"
      SQL = SQL & " and cidade = '" & TabCEP.Fields("cidade").Value & "'"
      SQL = SQL & " and uf = '" & TabCEP.Fields("uf").Value & "'"
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabConsulta.EOF Then
         SQL = "insert into CEP_VACA "
         SQL = SQL & " values("
            SQL = SQL & "'" & Trim(TabCEP.Fields("cep_id").Value) & "'"
            SQL = SQL & ",'" & TabCEP.Fields("cidade").Value & "'"
            SQL = SQL & ",'" & TabCEP.Fields("uf").Value & "'"
            SQL = SQL & ",0" & TabCEP.Fields("IBGE_ID").Value
         SQL = SQL & " )"
         CONECTA_RETAGUARDA.Execute SQL
         cmdTabela.Caption = "incluindo " & Trim(TabCEP.Fields("cep_id").Value) & "   " & TabCEP.Fields("cidade").Value
         cmdTabela.Refresh
         Me.Caption = "" & TabCEP.Fields("IBGE_ID").Value
         Else
            If Len(TabCEP.Fields("IBGE_ID").Value) > 5 Then
               SQL = "update CEP_VACA set "
               SQL = SQL & " IBGE_ID = " & TabCEP.Fields("IBGE_ID").Value
               SQL = SQL & " where cep_ID = '" & TabCEP.Fields("cep_id").Value & "'"
               SQL = SQL & " and cidade = '" & TabCEP.Fields("cidade").Value & "'"
               SQL = SQL & " and uf = '" & TabCEP.Fields("uf").Value & "'"
               CONECTA_RETAGUARDA.Execute SQL
         cmdCaixa.Caption = "alterando" & Trim(TabCEP.Fields("cep_id").Value) & "   " & TabCEP.Fields("cidade").Value
         cmdCaixa.Refresh
         Me.Caption = "" & TabCEP.Fields("IBGE_ID").Value
            End If
      End If

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      DoEvents

      TabCEP.MoveNext
   Wend

   If TabCEP.State = 1 Then _
      TabCEP.Close

MsgBox "OK"
End
End Sub

Private Sub Command18_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PROGRAMA", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "GERA_IMPRESSAO", "PROGRAMA") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PROGRAMA ADD GERA_IMPRESSAO BIT"
         CONECTA_RETAGUARDA.Execute "UPDATE PROGRAMA SET GERA_IMPRESSAO = 'FALSE'"
      End If
   End If
   MsgBox "OK"
End Sub

Private Sub Command2_Click()
   SQL = "update produto set qtde = 0"
   CONECTA_RETAGUARDA.Execute SQL

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select NOTAENTRADAITEM.ENTRADA_ID, NOTAENTRADAITEM.SEQ_id, "
   SQL = SQL & " NOTAENTRADAITEM.PRODUTO_ID, "
   SQL = SQL & " NOTAENTRADAITEM.qtde_entrada , NOTAENTRADA.Status"
   SQL = SQL & " from NOTAENTRADA "
   SQL = SQL & " INNER JOIN NOTAENTRADAITEM "
   SQL = SQL & " ON NOTAENTRADA.ENTRADA_ID = NOTAENTRADAITEM.ENTRADA_ID"
   SQL = SQL & " where NOTAENTRADA.status = 'E'"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      SQL = "update produto set"
      SQL = SQL & " qtde = qtde + " & tpMOEDA(TabConsulta.Fields("qtde_entrada").Value)
      SQL = SQL & " where produto_id = " & TabConsulta.Fields("produtO_id").Value
      CONECTA_RETAGUARDA.Execute SQL

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from INVENTARIO "
   SQL = SQL & " where status = 'F'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      If Trim(TabConsulta.Fields("tipo_mov").Value) = "E" Then
         SQL = "update produto set"
         SQL = SQL & " qtde = qtde + " & tpMOEDA(TabConsulta.Fields("qtd_PRIMEIRA").Value)
         SQL = SQL & " where produto_id = " & TabConsulta.Fields("produtO_id").Value
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If Trim(TabConsulta.Fields("tipo_mov").Value) = "S" Then
         SQL = "update produto set"
         SQL = SQL & " qtde = qtde - " & tpMOEDA(TabConsulta.Fields("qtd_PRIMEIRA").Value)
         SQL = SQL & " where produto_id = " & TabConsulta.Fields("produtO_id").Value
         CONECTA_RETAGUARDA.Execute SQL
      End If

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select PEDIDO.STATUS, PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.SEQ_ID, "
   SQL = SQL & " PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA"
   SQL = SQL & " from PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID"
   SQL = SQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      SQL = "update produto set"
      SQL = SQL & " qtde = qtde - " & tpMOEDA(TabConsulta.Fields("qtd_pedida").Value)
      SQL = SQL & " where produto_id = " & TabConsulta.Fields("produtO_id").Value
      CONECTA_RETAGUARDA.Execute SQL

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
MsgBox " e ai"
End Sub

Private Sub Command20_Click()
On Error Resume Next

   Msg = "Deseja Fazer Realmente esta operacao? EXCLUIR TODO BANCO DE DADOS"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      Msg = "ATENÇÃO TODAS INFORMAÇÕES DO BANCO DE DADOS SERÃO EXCLUIDAS, CONFIRMA?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then

         CONECTA_RETAGUARDA.Execute "DELETE from ABCREL"
         CONECTA_RETAGUARDA.Execute "DELETE from BOLETO"
         CONECTA_RETAGUARDA.Execute "DELETE from CABCOMPR"
         CONECTA_RETAGUARDA.Execute "DELETE from CABDEVENT"
         CONECTA_RETAGUARDA.Execute "DELETE from CABDEVSAI"
         CONECTA_RETAGUARDA.Execute "DELETE from CARTAOBARRA"
         CONECTA_RETAGUARDA.Execute "DELETE from cheque"
         CONECTA_RETAGUARDA.Execute "DELETE from CONTROLEPERDAITEM"
         CONECTA_RETAGUARDA.Execute "DELETE from CONTROLEPERDA"
         CONECTA_RETAGUARDA.Execute "DELETE from Endereco"
         CONECTA_RETAGUARDA.Execute "DELETE from IMPREL"
         CONECTA_RETAGUARDA.Execute "DELETE from IMPRESSORA"
         CONECTA_RETAGUARDA.Execute "DELETE from Indice"

         CONECTA_RETAGUARDA.Execute "DELETE from PEDIDOITEM"
         CONECTA_RETAGUARDA.Execute "DELETE from CUPOM"
         CONECTA_RETAGUARDA.Execute "DELETE from ENTREGA"
         CONECTA_RETAGUARDA.Execute "DELETE from PEDIDO"
         CONECTA_RETAGUARDA.Execute "DELETE from CAIXADIAITEM"
         CONECTA_RETAGUARDA.Execute "DELETE from CAIXADIA"
         CONECTA_RETAGUARDA.Execute "DELETE from CAIXATESORARIAitem"
         CONECTA_RETAGUARDA.Execute "DELETE from CAIXATESORARIA"
         CONECTA_RETAGUARDA.Execute "DELETE from CHEQUE"
         CONECTA_RETAGUARDA.Execute "DELETE from ITEMLANCAMENTO"
         CONECTA_RETAGUARDA.Execute "DELETE from ENDERECO"
         CONECTA_RETAGUARDA.Execute "DELETE from ERRO"
         CONECTA_RETAGUARDA.Execute "DELETE from FONE"
         CONECTA_RETAGUARDA.Execute "DELETE from FORNECEDOR"
         CONECTA_RETAGUARDA.Execute "DELETE from IE"
         CONECTA_RETAGUARDA.Execute "DELETE from IM"
         CONECTA_RETAGUARDA.Execute "DELETE from NOTAENTRADAITEM"
         'CONECTA_RETAGUARDA.Execute "DELETE from ITEMRECEBIMENTO"
         CONECTA_RETAGUARDA.Execute "DELETE from LANCAMENTO"
         CONECTA_RETAGUARDA.Execute "DELETE from NFITEM"
         CONECTA_RETAGUARDA.Execute "DELETE from NF"
         CONECTA_RETAGUARDA.Execute "DELETE from NOTAENTRADA"
         CONECTA_RETAGUARDA.Execute "DELETE from CLIENTE"
         CONECTA_RETAGUARDA.Execute "DELETE from COMISSAOITEM"
         CONECTA_RETAGUARDA.Execute "DELETE from COMISSAO"
         CONECTA_RETAGUARDA.Execute "DELETE from CONTA"
         CONECTA_RETAGUARDA.Execute "DELETE from CONTACORRENTE"
         CONECTA_RETAGUARDA.Execute "DELETE from CUPOM"
         CONECTA_RETAGUARDA.Execute "DELETE from EMAIL"
         CONECTA_RETAGUARDA.Execute "DELETE from ESTABELECIMENTOACESSO WHERE USUARIO_ID <> 144"
         'CONECTA_RETAGUARDA.Execute "DELETE from EMPRESAPARAMETRO"
         CONECTA_RETAGUARDA.Execute "DELETE from VENDEDOR WHERE VENDEDOR_ID <> 0"
         CONECTA_RETAGUARDA.Execute "DELETE from EQUIPE WHERE EQUIPE_ID <> 1"
         CONECTA_RETAGUARDA.Execute "DELETE from FAMILIAPRODUTO"
         
         CONECTA_RETAGUARDA.Execute "DELETE from OSPECA"
         CONECTA_RETAGUARDA.Execute "DELETE from OSSERVICO"
         CONECTA_RETAGUARDA.Execute "DELETE from OSTAREFA"
         CONECTA_RETAGUARDA.Execute "DELETE from OSVEICULO"
         CONECTA_RETAGUARDA.Execute "DELETE from OS"
         CONECTA_RETAGUARDA.Execute "DELETE from OSEQUIPAMENTO"
         CONECTA_RETAGUARDA.Execute "DELETE from INVENTARIO"
         CONECTA_RETAGUARDA.Execute "DELETE from PRODUTO"
         CONECTA_RETAGUARDA.Execute "DELETE from IMPRESSORA"
         CONECTA_RETAGUARDA.Execute "DELETE from INDICE"
         'CONECTA_RETAGUARDA.Execute "DELETE from ITEMRECEBIMENTO"
         CONECTA_RETAGUARDA.Execute "DELETE from OBS"
         CONECTA_RETAGUARDA.Execute "DELETE from RG"
         CONECTA_RETAGUARDA.Execute "DELETE from PERMISSAO WHERE USUID <> 144"
         CONECTA_RETAGUARDA.Execute "DELETE from USUARIO WHERE USUARIO_ID <> 144"
         CONECTA_RETAGUARDA.Execute "DELETE from TRANSPORTADORA"
MsgBox "banco MEGASIM zerado, ok"
      End If
   End If
End Sub

Private Sub Command22_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUTOLOTE", "U") = False Then
      SQL = "CREATE TABLE [dbo].[PRODUTOLOTE]("
      SQL = SQL & " [PRODUTOLOTE_ID] [bigint] NOT NULL,"
      SQL = SQL & " [EMPRESA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [BIGINT] NOT NULL,"
      SQL = SQL & " [NUMR_LOTE] [NVARCHAR](max) NOT NULL,"
      SQL = SQL & " [DATA_CAD] [datetime] NOT NULL,"
      SQL = SQL & " [DATA_FAB] [datetime] NOT NULL,"
      SQL = SQL & " [DATA_VENC] [datetime] NOT NULL,"
      SQL = SQL & " [SITUACAO] [NVARCHAR](1) NOT NULL,"
      SQL = SQL & " [ENTRADA_ID] [BIGINT] ,"
      SQL = SQL & " CONSTRAINT [PK_PRODUTOLOTE] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([PRODUTOLOTE_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_PRODUTOLOTE", "") = False Then
      SQL = "ALTER TABLE PRODUTOLOTE ADD CONSTRAINT PK_PRODUTOLOTE PRIMARY KEY (PRODUTOLOTE_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PRODUTOLOTE_PRODUTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[PRODUTOLOTE]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTOLOTE_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTOLOTE] CHECK CONSTRAINT [FK_PRODUTOLOTE_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub cmdCOMISSAO_Click()
   ATUALIZA_TABELA_COMISSAO
   MsgBox "ok"
End Sub

Private Sub Command26_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "TROCAPRODUTO", "") = True Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from TROCAPRODUTO "
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTemp.EOF Then
         If Not IsNull(TabTemp.Fields(0).Value) Then
            If TabTemp.Fields(0).Value >= 1 Then
               MsgBox "Verificar tabela TROCAPRODUTO "
               Exit Sub
            End If
         End If
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      CONECTA_RETAGUARDA.Execute "DROP TABLE [dbo].[TROCAPRODUTO]"
   End If

   SQL = "CREATE TABLE [dbo].[TROCAPRODUTO]("
   SQL = SQL & " [TROCA_ID] [bigint] NOT NULL,"
   SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
   SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
   SQL = SQL & " [PEDIDO_ID_TROCADO] [bigint] NOT NULL,"
   SQL = SQL & " [USUARIO_ID] [bigint] NOT NULL,"
   SQL = SQL & " [QTDE] [float] NOT NULL,"
   SQL = SQL & " [CODG_FUNC] [int] NOT NULL,"
   SQL = SQL & " [DT_TROCA] [datetime] NOT NULL,"
   SQL = SQL & " CONSTRAINT [PK_TROCAPRODUTO] PRIMARY KEY CLUSTERED"
   SQL = SQL & " ([TROCA_ID] Asc)"
   SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
   SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
   CONECTA_RETAGUARDA.Execute SQL


   SQL = " ALTER TABLE [dbo].[TROCAPRODUTO]  WITH CHECK ADD  CONSTRAINT [FK_TROCAPRODUTO_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
   SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = " ALTER TABLE [dbo].[TROCAPRODUTO] CHECK CONSTRAINT [FK_TROCAPRODUTO_PRODUTO]"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = " ALTER TABLE [dbo].[TROCAPRODUTO]  WITH CHECK ADD  CONSTRAINT [FK_TROCAPRODUTO_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
   SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = " ALTER TABLE [dbo].[TROCAPRODUTO] CHECK CONSTRAINT [FK_TROCAPRODUTO_PEDIDO]"
   CONECTA_RETAGUARDA.Execute SQL


   SQL = " ALTER TABLE [dbo].[TROCAPRODUTO]  WITH CHECK ADD  CONSTRAINT [FK_TROCAPRODUTO_PEDIDO_TROCA] FOREIGN KEY([PEDIDO_ID_TROCADO])"
   SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = " ALTER TABLE [dbo].[TROCAPRODUTO] CHECK CONSTRAINT [FK_TROCAPRODUTO_PEDIDO_TROCA]"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CLIENTESALDO", "") = False Then
      SQL = "CREATE TABLE [dbo].[CLIENTESALDO]("
      SQL = SQL & " [CLIENTESALDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CLIENTE_ID] [bigint] NOT NULL,"
      SQL = SQL & " [VALORSALDO] [float] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_CLIENTESALDO] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([CLIENTESALDO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) "
      SQL = SQL & " ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[CLIENTESALDO]  WITH CHECK ADD  CONSTRAINT [FK_CLIENTESALDO_CLIENTE] FOREIGN KEY([CLIENTE_ID])"
      SQL = SQL & " References [dbo].[Cliente]([CLIENTE_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[CLIENTESALDO] CHECK CONSTRAINT [FK_CLIENTESALDO_CLIENTE]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "OK"
End Sub

Private Sub Command27_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "IMPRESSORA", "U") = True Then
      MsgBox " CRIAR CHAVE MANUALMENTE (CAIXA_IMRPESSORA)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_IMPRESSORA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IMPRESSORA ADD CONSTRAINT PK_IMPRESSORA PRIMARY KEY (IMPRESSORA_ID)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "IMPRESSORA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE IMPRESSORA ADD EMPRESA_ID INT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_IMPRESSORA_EMPRESA", "") = False Then
         SQL = " alter table IMPRESSORA "
         SQL = SQL & " add constraINT FK_IMPRESSORA_EMPRESA foreign key (EMPRESA_ID)"
         SQL = SQL & " References EMPRESA(EMPRESA_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      Else
         SQL = "CREATE TABLE [dbo].[IMPRESSORA]("
         SQL = SQL & " [IMPRESSORA_ID] [int] NOT NULL,"
         SQL = SQL & " [DESCRICAO] [nvarchar](50) NOT NULL,"
         SQL = SQL & " [NUMR_SERIE_IMP] [nvarchar](50) NOT NULL,"
         SQL = SQL & " [NUMR_CAIXA] [int] NOT NULL,"
         SQL = SQL & " [CONTA_REINICIO] [int] NOT NULL,"
         SQL = SQL & " CONSTRAINT [PK_IMPRESSORA] PRIMARY KEY CLUSTERED([IMPRESSORA_ID] Asc)"
         SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, "
         SQL = SQL & " ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
         'CONECTA_RETAGUARDA.Execute SQL
   End If
'==================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "INDICE", "U") = True Then
      MsgBox "ORGANIZAR TABELA INDICE"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_INDICE", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INDICE DROP CONSTRAINT IX_INDICE "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INDICE", "IX_INDICE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INDICE DROP CONSTRAINT IX_INDICE "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "INDICE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INDICE DROP COLUMN EMPRESA_ID"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_INDICE", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE INDICE ADD CONSTRAINT PK_INDICE PRIMARY KEY (INDICE_ID,IMPRESSORA_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_INDICE_IMPRESSORA", "") = False Then
         SQL = " alter table INDICE "
         SQL = SQL & " add constraINT FK_INDICE_IMPRESSORA foreign key (IMPRESSORA_ID)"
         SQL = SQL & " References IMPRESSORA(IMPRESSORA_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      Else
         SQL = " CREATE TABLE [dbo].[INDICE]("
         SQL = SQL & " [INDICE_ID] [int] NOT NULL,"
         SQL = SQL & " [IMPRESSORA_ID] [int] NOT NULL,"
         SQL = SQL & " [ALIQUOTA] [int] NOT NULL,"
         SQL = SQL & " CONSTRAINT [PK_INDICE] PRIMARY KEY CLUSTERED"
         SQL = SQL & " ([INDICE_ID] ASC,[IMPRESSORA_ID] Asc"
         SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
         SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
         SQL = SQL & " ) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[INDICE]  WITH CHECK ADD  CONSTRAINT [FK_INDICE_IMPRESSORA] FOREIGN KEY([IMPRESSORA_ID])"
         SQL = SQL & " References [dbo].[IMPRESSORA]([IMPRESSORA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[INDICE] CHECK CONSTRAINT [FK_INDICE_IMPRESSORA]"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   MsgBox "Ok  =  " & CONT_N
   SQL = ""
End Sub

Private Sub CMDIMPREL_Click()
   CRIA_IMPREL
   MsgBox "ok"
End Sub

Private Sub Command4_Click()
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select pedido_id,cliente_id,nome_cliente from PEDIDO "
   SQL = SQL & "  WHERE nome_cliente IS NULL OR NOME_CLIENTE = ''"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      Command4.Caption = TabTemp.Fields("pedido_id").Value
      DoEvents
      SQL = "update PEDIDO set nome_cliente = 'Não Informado' "
      SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
      CONECTA_RETAGUARDA.Execute SQL

      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
   MsgBox "ok"
End Sub

Private Sub Command8_Click()
On Error Resume Next

   Msg = "Deseja Fazer Realmente esta operacao? EXCLUIR TODO BANCO DE DADOS"
   PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      Msg = "ATENÇÃO TODAS INFORMAÇÕES DO BANCO DE DADOS SERÃO EXCLUIDAS, CONFIRMA?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then

         If CONECTA_GLOBAL.State = 1 Then _
            CONECTA_GLOBAL.Close

         ABRE_BANCO_GLOBAL

         If CONECTA_GLOBAL.State <> 1 Then
            MsgBox "Banco GLOBAL não conectado."
            Exit Sub
         End If

         CONECTA_GLOBAL.Execute "DELETE from MFA010"
         CONECTA_GLOBAL.Execute "DELETE from MFAOBS"
         CONECTA_GLOBAL.Execute "DELETE from MFI010"
         CONECTA_GLOBAL.Execute "DELETE from MFT010"
         CONECTA_GLOBAL.Execute "DELETE from SA1010"
         CONECTA_GLOBAL.Execute "DELETE from SA2010"
         CONECTA_GLOBAL.Execute "DELETE from SB1010"
         CONECTA_GLOBAL.Execute "DELETE from SB2010"
         CONECTA_GLOBAL.Execute "DELETE from SE1010"
         CONECTA_GLOBAL.Execute "DELETE from TB_CARTACORRECAO_NFE"

MsgBox "banco GLOBAL zerado, ok"
      End If
   End If
End Sub

Private Sub Command9_Click()
   If Not EXISTE_OBJ_BANCO("RETAGUARDA", "CHAMADO", "U") = True Then
      SQL = "CREATE TABLE [dbo].[CHAMADO]("
      SQL = SQL & " [CHAMADO_ID] [nchar](10) NOT NULL,"
      SQL = SQL & " [PESSOA_ID] [nchar](10) NOT NULL,"
      SQL = SQL & " [HISTORICO] [text] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_CHAMADO] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ("
      SQL = SQL & " [CHAMADO_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & " ) "
      SQL = SQL & " ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
End Sub

Private Sub cmdKardex_Click()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "QryItensNotasFiscaisVenda", "") = True Then
      SQL = " ALTER VIEW [dbo].[QryItensNotasFiscaisVenda]"
      Else: SQL = " CREATE VIEW [dbo].[QryItensNotasFiscaisVenda]"
   End If
   SQL = SQL & " AS"
   SQL = SQL & " select PEDIDO.PEDIDO_ID, PEDIDOITEM.PRODUTO_ID, "
   SQL = SQL & " PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.DESCRICAO, PEDIDOITEM.VALOR_ITEM, "
   SQL = SQL & " PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.CFOP_id, PEDIDOITEM.STRIBUTARIA, PEDIDOITEM.VLRBASEICMS,"
   SQL = SQL & " PEDIDOITEM.PERCICMS, PEDIDOITEM.VLRICMS, PEDIDOITEM.PERCICMSSUBST, "
   SQL = SQL & " PEDIDOITEM.VLRICMSSUBST, PEDIDOITEM.PERCREDUCAOICMS, PEDIDOITEM.PERCIVA, "
   SQL = SQL & " CLIENTE.CLIENTE_ID, CLIENTE.NOME, PEDIDO.DT_REQ,"
   SQL = SQL & " PEDIDOITEM.Status"
   SQL = SQL & " from PRODUTO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PRODUTO.PRODUTO_ID = PEDIDOITEM.PRODUTO_ID "
   SQL = SQL & " INNER JOIN PEDIDO "
   SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = PEDIDO.PEDIDO_ID "
   SQL = SQL & " INNER JOIN CLIENTE "
   SQL = SQL & " ON PEDIDO.CLIENTE_ID = CLIENTE.CLIENTE_ID"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "QryNotaEntrada", "") = True Then
      SQL = " ALTER VIEW [dbo].[QryNotaEntrada]"
      Else: SQL = " CREATE VIEW [dbo].[QryNotaEntrada]"
   End If
   SQL = SQL & " AS "
   SQL = SQL & " select * from NOTAENTRADA"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "QryItensNotaEntrada", "") = True Then
      SQL = " ALTER VIEW [dbo].[QryItensNotaEntrada]"
      Else: SQL = " CREATE VIEW [dbo].[QryItensNotaEntrada]"
   End If
   SQL = SQL & " AS "
   SQL = SQL & " select NOTAENTRADAITEM.PRODUTO_ID,PRODUTO.DESCRICAO, "
   SQL = SQL & " NOTAENTRADAITEM.PRECO_CUSTO, NOTAENTRADAITEM.qtde_entrada, "
   SQL = SQL & " NOTAENTRADAITEM.ENTRADA_ID, QryNOTAENTRADA.pedidocompra_id, QryNotaEntrada.NUMR_NOTA,"
   SQL = SQL & " NOTAENTRADAITEM.SEQ_id, QryNotaEntrada.FORNECEDOR_ID, "
   SQL = SQL & " FORNECEDOR.NOME, NOTAENTRADAITEM.CFOP_id, NOTAENTRADAITEM.PERC_IPI, NOTAENTRADAITEM.PERC_ICMS, "
   SQL = SQL & " NOTAENTRADAITEM.VALOR_DESCONTO, QryNotaEntrada.DT_ENTRADA,NOTAENTRADAITEM.Status"
   SQL = SQL & " from NOTAENTRADAITEM "
   SQL = SQL & " INNER JOIN QryNotaEntrada "
   SQL = SQL & " ON NOTAENTRADAITEM.ENTRADA_ID = QryNotaEntrada.ENTRADA_ID "
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON NOTAENTRADAITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
   SQL = SQL & " INNER JOIN FORNECEDOR "
   SQL = SQL & " ON QryNotaEntrada.FORNECEDOR_ID = FORNECEDOR.FORNECEDOR_ID"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "QryFinalKardex", "") = True Then
      SQL = "ALTER VIEW [dbo].[QryFinalKardex]"
      Else: SQL = " CREATE VIEW [dbo].[QryFinalKardex]"
   End If
   SQL = SQL & " AS "
   SQL = SQL & " (select QryItensNotaEntrada.CODG_PRODUTO, QryItensNotaEntrada.DESCRICAO, "
   SQL = SQL & " QryItensNotaEntrada.qtde_entrada, QryItensNotaEntrada.PRECO_CUSTO, "
   SQL = SQL & " QryItensNotaEntrada.fornecedor_id, QryItensNotaEntrada.NOME, QryItensNotaEntrada.DT_ENTRADA, "
   SQL = SQL & " QryItensNotaEntrada.NUMR_NOTA, QryItensNotaEntrada.Qtde, "
   SQL = SQL & " QryItensNotaEntrada.Status from QryItensNotaEntrada) "
   SQL = SQL & " Union  (select QryItensNotasFiscaisVenda.CODG_PRODUTO, QryItensNotasFiscaisVenda.DESCRICAO, "
   SQL = SQL & " QryItensNotasFiscaisVenda.QTD_PEDIDA, QryItensNotasFiscaisVenda.VALOR_ITEM, QryItensNotasFiscaisVenda.CLIENTE_ID, "
   SQL = SQL & " QryItensNotasFiscaisVenda.NOME, QryItensNotasFiscaisVenda.DT_REQ, QryItensNotasFiscaisVenda.PEDIDO_ID, "
   SQL = SQL & " QryItensNotasFiscaisVenda.QTDE , QryItensNotasFiscaisVenda.Status "
   SQL = SQL & " from QryItensNotasFiscaisVenda)"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_KARDEX", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_KARDEX"
   
   SQL = "CREATE PROCEDURE [dbo].[SP_KARDEX] (@Dt_Inicial datetime, @Dt_Final datetime, @tipo varchar(50), @Codigo int, @Produto varchar(50))   as "
   SQL = SQL & " SET NOCOUNT on"
   SQL = SQL & " DECLARE @Tipo1 varchar(50);"
   SQL = SQL & " DECLARE @Codigo1 int;"
   SQL = SQL & " DECLARE @Produto1 varchar(50);"
   SQL = SQL & " set @tipo1 = @tipo"
   SQL = SQL & " set @Codigo1 = @Codigo"
   SQL = SQL & " set @Produto1 = @Produto"
   SQL = SQL & " begin"
   SQL = SQL & " if @Tipo1 = ''"
   SQL = SQL & " begin"
   SQL = SQL & " if @Produto1 = ''"
   SQL = SQL & " select 'ENTRADA' as ENTRADA, QryItensNotaEntrada.CODG_PRODUTO, QryItensNotaEntrada.DESCRICAO, QryItensNotaEntrada.qtde_entrada, QryItensNotaEntrada.PRECO_CUSTO,"
   SQL = SQL & " QryItensNotaEntrada.fornecedor_id, QryItensNotaEntrada.NOME, QryItensNotaEntrada.DT_ENTRADA, QryItensNotaEntrada.NUMR_NOTA,"
   SQL = SQL & " QryItensNotaEntrada.Qtde"
   SQL = SQL & " from QryItensNotaEntrada Where QryItensNotaEntrada.DT_ENTRADA >= @Dt_Inicial and QryItensNotaEntrada.DT_ENTRADA <= @Dt_Final"
   SQL = SQL & " Union"
   SQL = SQL & " select 'SAIDA' as SAIDA, QryItensNotasFiscaisVenda.CODG_PRODUTO, QryItensNotasFiscaisVenda.DESCRICAO, QryItensNotasFiscaisVenda.QTD_PEDIDA,"
   SQL = SQL & " QryItensNotasFiscaisVenda.VALOR_ITEM, QryItensNotasFiscaisVenda.CLIENTE_ID, QryItensNotasFiscaisVenda.NOME,"
   SQL = SQL & " QryItensNotasFiscaisVenda.DT_REQ , QryItensNotasFiscaisVenda.PEDIDO_ID, QryItensNotasFiscaisVenda.Qtde"
   SQL = SQL & " from QryItensNotasFiscaisVenda Where QryItensNotasFiscaisVenda.DT_REQ >= @Dt_Inicial and QryItensNotasFiscaisVenda.DT_REQ <= @Dt_Final"
   SQL = SQL & " order by QryItensNotaEntrada.DT_ENTRADA"
   SQL = SQL & " Else"
   SQL = SQL & " select 'ENTRADA' as ENTRADA, QryItensNotaEntrada.CODG_PRODUTO, QryItensNotaEntrada.DESCRICAO, QryItensNotaEntrada.qtde_entrada, QryItensNotaEntrada.PRECO_CUSTO,"
   SQL = SQL & " QryItensNotaEntrada.fornecedor_id, QryItensNotaEntrada.NOME, QryItensNotaEntrada.DT_ENTRADA, QryItensNotaEntrada.NUMR_NOTA,"
   SQL = SQL & " QryItensNotaEntrada.Qtde"
   SQL = SQL & " from QryItensNotaEntrada Where QryItensNotaEntrada.DT_ENTRADA >= @Dt_Inicial and QryItensNotaEntrada.DT_ENTRADA <= @Dt_Final and QryItensNotaEntrada.CODG_PRODUTO = @Produto"
   SQL = SQL & " Union "
   SQL = SQL & " select 'SAIDA' as SAIDA, QryItensNotasFiscaisVenda.CODG_PRODUTO, QryItensNotasFiscaisVenda.DESCRICAO, QryItensNotasFiscaisVenda.QTD_PEDIDA,"
   SQL = SQL & " QryItensNotasFiscaisVenda.VALOR_ITEM, QryItensNotasFiscaisVenda.CLIENTE_ID, QryItensNotasFiscaisVenda.NOME,"
   SQL = SQL & " QryItensNotasFiscaisVenda.DT_REQ , QryItensNotasFiscaisVenda.PEDIDO_ID, QryItensNotasFiscaisVenda.Qtde"
   SQL = SQL & " from QryItensNotasFiscaisVenda Where QryItensNotasFiscaisVenda.DT_REQ >= @Dt_Inicial and QryItensNotasFiscaisVenda.DT_REQ <= @Dt_Final and QryItensNotasFiscaisVenda.CODG_PRODUTO = @Produto"
   SQL = SQL & " order by QryItensNotaEntrada.DT_ENTRADA"
   SQL = SQL & " End"
   SQL = SQL & " Else"
   SQL = SQL & " begin"
   SQL = SQL & " if @Produto1 = ''"
   SQL = SQL & " begin"
   SQL = SQL & " if @Codigo1 = 0"
   SQL = SQL & " select 'ENTRADA' as ENTRADA, QryItensNotaEntrada.CODG_PRODUTO, QryItensNotaEntrada.DESCRICAO, QryItensNotaEntrada.qtde_entrada, QryItensNotaEntrada.PRECO_CUSTO,"
   SQL = SQL & " QryItensNotaEntrada.fornecedor_id, QryItensNotaEntrada.NOME, QryItensNotaEntrada.DT_ENTRADA, QryItensNotaEntrada.NUMR_NOTA,"
   SQL = SQL & " QryItensNotaEntrada.Qtde"
   SQL = SQL & " from QryItensNotaEntrada Where QryItensNotaEntrada.DT_ENTRADA >= @Dt_Inicial and QryItensNotaEntrada.DT_ENTRADA <= @Dt_Final and 'ENTRADA' = @tipo"
   SQL = SQL & " Union"
   SQL = SQL & " select 'SAIDA' as SAIDA, QryItensNotasFiscaisVenda.CODG_PRODUTO, QryItensNotasFiscaisVenda.DESCRICAO, QryItensNotasFiscaisVenda.QTD_PEDIDA,"
   SQL = SQL & " QryItensNotasFiscaisVenda.VALOR_ITEM, QryItensNotasFiscaisVenda.CLIENTE_ID, QryItensNotasFiscaisVenda.NOME,"
   SQL = SQL & " QryItensNotasFiscaisVenda.DT_REQ , QryItensNotasFiscaisVenda.PEDIDO_ID, QryItensNotasFiscaisVenda.Qtde"
   SQL = SQL & " from QryItensNotasFiscaisVenda Where QryItensNotasFiscaisVenda.DT_REQ >= @Dt_Inicial and QryItensNotasFiscaisVenda.DT_REQ <= @Dt_Final and 'SAIDA' = @tipo"
   SQL = SQL & " order by QryItensNotaEntrada.DT_ENTRADA"
   SQL = SQL & " Else"
   SQL = SQL & " select 'ENTRADA' as ENTRADA, QryItensNotaEntrada.CODG_PRODUTO, QryItensNotaEntrada.DESCRICAO, QryItensNotaEntrada.qtde_entrada, QryItensNotaEntrada.PRECO_CUSTO,"
   SQL = SQL & " QryItensNotaEntrada.fornecedor_id, QryItensNotaEntrada.NOME, QryItensNotaEntrada.DT_ENTRADA, QryItensNotaEntrada.NUMR_NOTA,"
   SQL = SQL & " QryItensNotaEntrada.Qtde"
   SQL = SQL & " from QryItensNotaEntrada Where QryItensNotaEntrada.DT_ENTRADA >= @Dt_Inicial and QryItensNotaEntrada.DT_ENTRADA <= @Dt_Final and 'ENTRADA' = @tipo and QryItensNotaEntrada.fornecedor_id = @Codigo"
   SQL = SQL & " Union"
   SQL = SQL & " select 'SAIDA' as SAIDA, QryItensNotasFiscaisVenda.CODG_PRODUTO, QryItensNotasFiscaisVenda.DESCRICAO, QryItensNotasFiscaisVenda.QTD_PEDIDA,"
   SQL = SQL & " QryItensNotasFiscaisVenda.VALOR_ITEM, QryItensNotasFiscaisVenda.CLIENTE_ID, QryItensNotasFiscaisVenda.NOME,"
   SQL = SQL & " QryItensNotasFiscaisVenda.DT_REQ , QryItensNotasFiscaisVenda.PEDIDO_ID, QryItensNotasFiscaisVenda.Qtde"
   SQL = SQL & " from QryItensNotasFiscaisVenda Where QryItensNotasFiscaisVenda.DT_REQ >= @Dt_Inicial and QryItensNotasFiscaisVenda.DT_REQ <= @Dt_Final and 'SAIDA' = @tipo and QryItensNotasFiscaisVenda.CLIENTE_ID = @Codigo"
   SQL = SQL & " order by QryItensNotaEntrada.DT_ENTRADA"
   SQL = SQL & " End"
   SQL = SQL & " Else"
   SQL = SQL & " begin"
   SQL = SQL & " if @Codigo1 = 0"
   SQL = SQL & " select 'ENTRADA' as ENTRADA, QryItensNotaEntrada.CODG_PRODUTO, QryItensNotaEntrada.DESCRICAO, QryItensNotaEntrada.qtde_entrada, QryItensNotaEntrada.PRECO_CUSTO,"
   SQL = SQL & " QryItensNotaEntrada.fornecedor_id, QryItensNotaEntrada.NOME, QryItensNotaEntrada.DT_ENTRADA, QryItensNotaEntrada.NUMR_NOTA,"
   SQL = SQL & " QryItensNotaEntrada.Qtde "
   SQL = SQL & " from QryItensNotaEntrada Where QryItensNotaEntrada.DT_ENTRADA >= @Dt_Inicial and QryItensNotaEntrada.DT_ENTRADA <= @Dt_Final and 'ENTRADA' = @tipo and QryItensNotaEntrada.CODG_PRODUTO = @Produto"
   SQL = SQL & " Union"
   SQL = SQL & " select 'SAIDA' as SAIDA, QryItensNotasFiscaisVenda.CODG_PRODUTO, QryItensNotasFiscaisVenda.DESCRICAO, QryItensNotasFiscaisVenda.QTD_PEDIDA,"
   SQL = SQL & " QryItensNotasFiscaisVenda.VALOR_ITEM, QryItensNotasFiscaisVenda.CLIENTE_ID, QryItensNotasFiscaisVenda.NOME,"
   SQL = SQL & " QryItensNotasFiscaisVenda.DT_REQ , QryItensNotasFiscaisVenda.PEDIDO_ID, QryItensNotasFiscaisVenda.Qtde"
   SQL = SQL & " from QryItensNotasFiscaisVenda Where QryItensNotasFiscaisVenda.DT_REQ >= @Dt_Inicial and QryItensNotasFiscaisVenda.DT_REQ <= @Dt_Final and 'SAIDA' = @tipo and QryItensNotasFiscaisVenda.CODG_PRODUTO = @Produto"
   SQL = SQL & " order by QryItensNotaEntrada.DT_ENTRADA"
   SQL = SQL & " Else"
   SQL = SQL & " select 'ENTRADA' as ENTRADA, QryItensNotaEntrada.CODG_PRODUTO, QryItensNotaEntrada.DESCRICAO, QryItensNotaEntrada.qtde_entrada, QryItensNotaEntrada.PRECO_CUSTO,"
   SQL = SQL & " QryItensNotaEntrada.fornecedor_id, QryItensNotaEntrada.NOME, QryItensNotaEntrada.DT_ENTRADA, QryItensNotaEntrada.NUMR_NOTA,"
   SQL = SQL & " QryItensNotaEntrada.Qtde "
   SQL = SQL & " from QryItensNotaEntrada Where QryItensNotaEntrada.DT_ENTRADA >= @Dt_Inicial and QryItensNotaEntrada.DT_ENTRADA <= @Dt_Final and 'ENTRADA' = @tipo and QryItensNotaEntrada.fornecedor_id = @Codigo and QryItensNotaEntrada.CODG_PRODUTO = @Produto"
   SQL = SQL & " Union"
   SQL = SQL & " select 'SAIDA' as SAIDA, QryItensNotasFiscaisVenda.CODG_PRODUTO, QryItensNotasFiscaisVenda.DESCRICAO, QryItensNotasFiscaisVenda.QTD_PEDIDA,"
   SQL = SQL & " QryItensNotasFiscaisVenda.VALOR_ITEM, QryItensNotasFiscaisVenda.CLIENTE_ID, QryItensNotasFiscaisVenda.NOME,"
   SQL = SQL & " QryItensNotasFiscaisVenda.DT_REQ , QryItensNotasFiscaisVenda.PEDIDO_ID, QryItensNotasFiscaisVenda.Qtde"
   SQL = SQL & " from QryItensNotasFiscaisVenda Where QryItensNotasFiscaisVenda.DT_REQ >= @Dt_Inicial and QryItensNotasFiscaisVenda.DT_REQ <= @Dt_Final and 'SAIDA' = @tipo and QryItensNotasFiscaisVenda.CLIENTE_ID = @Codigo  and QryItensNotasFiscaisVenda.CODG_PRODUTO = @Produto"
   SQL = SQL & " order by QryItensNotaEntrada.DT_ENTRADA"
   SQL = SQL & " End"
   SQL = SQL & " End"
   SQL = SQL & " End"
   CONECTA_RETAGUARDA.Execute SQL

   MsgBox "Ok  =  " & CONT_N
End Sub

Private Sub Command6_Click()

SQL = "update ITEMLANCAMENTO set status = 'B'"
SQL = SQL & " Where ITEMLANCAMENTO.FORMAPAGTO_ID = 1"
SQL = SQL & " and ITEMLANCAMENTO.status = 'A'"
CONECTA_RETAGUARDA.Execute SQL

SQL = "update ITEMLANCAMENTO set dt_baixa = '25/05/2015'"
SQL = SQL & " Where ITEMLANCAMENTO.status = 'B' and ITEMLANCAMENTO.DT_BAIXA is null"
CONECTA_RETAGUARDA.Execute SQL

MsgBox "ok"

Exit Sub

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select ITEMLANCAMENTO.*, LANCAMENTO.dt_cad from LANCAMENTO "
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
   SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
   SQL = SQL & " order by ITEMLANCAMENTO.lancamento_id"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      If Not IsNull(TabConsulta!DT_BAIXA) Then
         If TabConsulta.Fields("dt_baixa").Value >= TabConsulta.Fields("dt_cad").Value Then
            SQL = "UPDATE ITEMLANCAMENTO SET "
               SQL = SQL & " Status = 'B'"
               'SQL = SQL & ", DT_baixa = '" & DMA(TabConsulta.Fields("dt_vencimento").Value) & "'"
            SQL = SQL & " Where Lancamento_id = " & TabConsulta.Fields("lancamento_id").Value
            SQL = SQL & " and NUMR_DOC = " & TabConsulta.Fields("NUMR_DOC").Value
            SQL = SQL & " and SEQ = " & TabConsulta.Fields("SEQ").Value
            CONECTA_RETAGUARDA.Execute SQL
         End If
         Else
            SQL = "UPDATE ITEMLANCAMENTO SET "
               SQL = SQL & " Status = 'A'"
               SQL = SQL & ", DT_baixa = null"
            SQL = SQL & " Where Lancamento_id = " & TabConsulta.Fields("lancamento_id").Value
            SQL = SQL & " and NUMR_DOC = " & TabConsulta.Fields("NUMR_DOC").Value
            SQL = SQL & " and SEQ = " & TabConsulta.Fields("SEQ").Value
            CONECTA_RETAGUARDA.Execute SQL
      End If
      'If TabConsulta.Fields("formapagto_id").Value = 1 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 2 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 3 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 4 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 6 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 7 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 8 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 9 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 10 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 11 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 12 Or _
      '   TabConsulta.Fields("formapagto_id").Value = 13 Then
   
      '   If IsNull(TabConsulta!DT_BAIXA) Then
      '      SQL = "UPDATE ITEMLANCAMENTO SET "
      '      SQL = SQL & " Status = 'B'"
      '      SQL = SQL & ", DT_BAIXA = '" & now & "'"
      '      SQL = SQL & ", CODG_USU_BAIXA = " & usuario_id_N
      '      SQL = SQL & " Where Lancamento_id = " & TabConsulta.Fields("lancamento_id").Value
      '      SQL = SQL & " and NUMR_DOC = " & TabConsulta.Fields("NUMR_DOC").Value
      '      SQL = SQL & " and SEQ = " & TabConsulta.Fields("SEQ").Value
      '      CONECTA_RETAGUARDA.Execute SQL
      '      Else
      '         If Not IsDate(TabConsulta!DT_BAIXA) Then
      '            SQL = "UPDATE ITEMLANCAMENTO SET "
      '            SQL = SQL & " Status = 'B'"
      '            SQL = SQL & ", DT_BAIXA = '" & now & "'"
      '            SQL = SQL & ", CODG_USU_BAIXA = " & usuario_id_N
      '            SQL = SQL & " Where Lancamento_id = " & TabConsulta.Fields("lancamento_id").Value
      '            SQL = SQL & " and NUMR_DOC = " & TabConsulta.Fields("NUMR_DOC").Value
      '            SQL = SQL & " and SEQ = " & TabConsulta.Fields("SEQ").Value
      '            CONECTA_RETAGUARDA.Execute SQL
      '         End If
      '   End If
      'End If

      'If Not IsNull(TabConsulta!DT_BAIXA) Then
      '   If IsDate(TabConsulta!DT_BAIXA) Then
      '      If Year(TabConsulta!DT_BAIXA) > 2000 Then
      '         SQL = "update ITEMLANCAMENTO set status = 'B' "
      '         SQL = SQL & " where status = 'A'"
      '         SQL = SQL & " and LANCAMENTO_ID = " & TabConsulta.Fields("lancamento_id").Value
      '         SQL = SQL & " and NUMR_DOC = " & TabConsulta.Fields("NUMR_DOC").Value
      '         SQL = SQL & " and SEQ = " & TabConsulta.Fields("SEQ").Value
      '         CONECTA_RETAGUARDA.Execute SQL
      '         Command6.Caption = TabConsulta.Fields("lancamento_id").Value
      '      End If
      '   End If
      'End If

      CONT_N = CONT_N + 1
      Command6.Caption = CONT_N

      TabConsulta.MoveNext
      DoEvents
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
MsgBox "Ok  =  " & CONT_N
End Sub

'===================================
'===================================
'===================================
Private Sub Command21_Click()

   CONECTA_RETAGUARDA.Execute "EXEC sp_resetstatus 'loja1'"

CONECTA_RETAGUARDA.Execute "ALTER DATABASE loja1 SET EMERGENCY"
CONECTA_RETAGUARDA.Execute "DBCC checkdb('loja1')"
CONECTA_RETAGUARDA.Execute "ALTER DATABASE loja1 SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
CONECTA_RETAGUARDA.Execute "DBCC CheckDB ('loja1', REPAIR_ALLOW_DATA_LOSS)"
CONECTA_RETAGUARDA.Execute "ALTER DATABASE loja1 SET MULTI_USER"

'GO

'-- Rebuild the index

'ALTER INDEX ALL ON [ TableName] REBUILD
End Sub

Private Sub Command25_Click()
'======================BANCO
   If EXISTE_OBJ_BANCO("RETAGUARDA", "BANCO", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "BANCO_ID", "BANCO") = False Then
         'CRIA CAMPO
         CONECTA_RETAGUARDA.Execute "ALTER TABLE BANCO ADD BANCO_ID BIGINT"

         'CRIA PROCEDURE ALIMENTAR SEQUENCIA CAMPO BANCO_ID
         If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_TABANCO", "") = True Then _
            CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_UPDATE_ID_TABANCO"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_TABANCO", "") = False Then
            SQL = "CREATE PROCEDURE SP_UPDATE_ID_TABANCO "
            SQL = SQL & " as "
            SQL = SQL & " DECLARE @Contador AS SMALLINT"
            SQL = SQL & " SET @Contador = 0"
            SQL = SQL & " Update Banco"
            SQL = SQL & " SET @Contador = @Contador + 1"
            SQL = SQL & " , BANCO_ID = @Contador"
            CONECTA_RETAGUARDA.Execute SQL
         End If

         CONECTA_RETAGUARDA.Execute "EXEC SP_UPDATE_ID_TABANCO "

         Alteração_Definição_Campo_Tabela "BANCO_ID", "BIGINT NOT NULL", "BANCO", "RETAGUARDA"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_BANCO", "") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE BANCO ADD CONSTRAINT PK_BANCO PRIMARY KEY (BANCO_ID)"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Banco", "BANCO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'BANCO.BANCO'" & "," & "'CODG_BANCO'" & "," & "'COLUMN'"
      
      Alteração_Definição_Campo_Tabela "CODG_BANCO", "BIGINT NOT NULL", "BANCO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Nome_Banco", "BANCO") = True Then _
         Alteração_Definição_Campo_Tabela "Nome_Banco", "nvarchar(100) NOT NULL", "BANCO", "RETAGUARDA"

      Else
         SQL = "CREATE TABLE [dbo].[BANCO]("
         SQL = SQL & " [BANCO_ID] [bigint] NOT NULL,"
         SQL = SQL & " [CODG_BANCO] [nvarchar](6) NOT NULL,"
         SQL = SQL & " [NOME_BANCO] [nvarchar](100) NOT NULL,"
         SQL = SQL & " CONSTRAINT [PK_BANCO] PRIMARY KEY CLUSTERED"
         SQL = SQL & " ([BANCO_ID] Asc)"
         SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
         SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) "
         SQL = SQL & " ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL
   End If

'======================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "AGENCIA", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "AGENCIA_ID", "AGENCIA") = False Then
         'CRIA CAMPO
         CONECTA_RETAGUARDA.Execute "ALTER TABLE AGENCIA ADD AGENCIA_ID BIGINT"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_TABAGENCIA", "") = True Then _
            CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_UPDATE_ID_TABAGENCIA"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_TABAGENCIA", "") = False Then
            SQL = "CREATE PROCEDURE SP_UPDATE_ID_TABAGENCIA "
            SQL = SQL & " as "
            SQL = SQL & " DECLARE @Contador AS SMALLINT"
            SQL = SQL & " SET @Contador = 0"
            SQL = SQL & " Update AGENCIA"
            SQL = SQL & " SET @Contador = @Contador + 1"
            SQL = SQL & " , AGENCIA_ID = @Contador"
            CONECTA_RETAGUARDA.Execute SQL
         End If

         CONECTA_RETAGUARDA.Execute "EXEC SP_UPDATE_ID_TABAGENCIA "

         Alteração_Definição_Campo_Tabela "AGENCIA_ID", "BIGINT NOT NULL", "AGENCIA", "RETAGUARDA"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_AGENCIA", "") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE AGENCIA ADD CONSTRAINT PK_AGENCIA PRIMARY KEY (AGENCIA_ID)"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_AGENCIA", "AGENCIA") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_AGENCIA", "nvarchar(10) NOT NULL", "AGENCIA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Banco", "AGENCIA") = True Then
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'AGENCIA.BANCO'" & "," & "'CODG_BANCO'" & "," & "'COLUMN'"
         Alteração_Definição_Campo_Tabela "CODG_BANCO", "BIGINT NOT NULL", "AGENCIA", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NOME_AGENCIA", "AGENCIA") = True Then _
         Alteração_Definição_Campo_Tabela "NOME_AGENCIA", "nvarchar(100) NOT NULL", "AGENCIA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "BANCO_ID", "AGENCIA") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE AGENCIA ADD BANCO_ID BIGINT"

         If TabAGENCIA.State = 1 Then _
            TabAGENCIA.Close

         SQL = "select * from AGENCIA "
         TabAGENCIA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabAGENCIA.EOF

            If TabBANCO.State = 1 Then _
               TabBANCO.Close

            SQL = "select banco_id from BANCO "
            SQL = SQL & " where codg_banco = '" & TabAGENCIA.Fields("codg_banco").Value & "'"
            TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabBANCO.EOF Then
               SQL = "update AGENCIA set "
               SQL = SQL & " banco_id = " & TabBANCO.Fields(0).Value
               SQL = SQL & " where codg_banco = " & TabAGENCIA.Fields("codg_banco").Value
               CONECTA_RETAGUARDA.Execute SQL
            End If
            If TabBANCO.State = 1 Then _
               TabBANCO.Close

            TabAGENCIA.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "BANCO_ID", "AGENCIA") = True Then _
         Alteração_Definição_Campo_Tabela "BANCO_ID", "BIGINT NOT NULL", "AGENCIA", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_AGENCIA_BANCO", "") = False Then
         SQL = "ALTER TABLE [dbo].[AGENCIA] WITH CHECK ADD CONSTRAINT [FK_AGENCIA_BANCO] FOREIGN KEY([BANCO_ID])"
         SQL = SQL & " References [dbo].[BANCO]([BANCO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[AGENCIA] CHECK CONSTRAINT [FK_AGENCIA_BANCO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENDERECO_ID", "AGENCIA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE AGENCIA ADD ENDERECO_ID BIGINT"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA","ENDERECO_ID", "AGENCIA") = True Then _
         Alteração_Definição_Campo_Tabela "ENDERECO_ID", "BIGINT", "AGENCIA", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_AGENCIA_ENDERECO", "") = False Then
         SQL = "ALTER TABLE [dbo].[AGENCIA] WITH CHECK ADD CONSTRAINT [FK_AGENCIA_ENDERECO] FOREIGN KEY([ENDERECO_ID])"
         SQL = SQL & " References [dbo].[ENDERECO]([ENDERECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[AGENCIA] CHECK CONSTRAINT [FK_AGENCIA_ENDERECO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      Else
         SQL = "CREATE TABLE [dbo].[AGENCIA]("
         SQL = SQL & " [AGENCIA_ID] [bigint] NOT NULL,"
         SQL = SQL & " [BANCO_ID] [bigint] NOT NULL,"
         SQL = SQL & " [ENDERECO_ID] [bigint],"
         SQL = SQL & " [NUMR_AGENCIA] [nvarchar](10) NOT NULL,"
         SQL = SQL & " [CODG_BANCO] [bigint] NOT NULL,"
         SQL = SQL & " [NOME_AGENCIA] [nvarchar](100) NOT NULL,"
         SQL = SQL & " CONSTRAINT [PK_AGENCIA] PRIMARY KEY CLUSTERED"
         SQL = SQL & " ([AGENCIA_ID] Asc)"
         SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
         SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
         SQL = SQL & " ) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[AGENCIA] WITH CHECK ADD CONSTRAINT [FK_AGENCIA_BANCO] FOREIGN KEY([BANCO_ID])"
         SQL = SQL & " References [dbo].[BANCO]([BANCO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[AGENCIA] CHECK CONSTRAINT [FK_AGENCIA_BANCO]"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[AGENCIA] WITH CHECK ADD CONSTRAINT [FK_AGENCIA_ENDERECO] FOREIGN KEY([ENDERECO_ID])"
         SQL = SQL & " References [dbo].[ENDERECO]([ENDERECO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[AGENCIA] CHECK CONSTRAINT [FK_AGENCIA_ENDERECO]"
         CONECTA_RETAGUARDA.Execute SQL
   End If

'======================CONTA
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CONTA", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTA_ID", "CONTA") = False Then
         'CRIA CAMPO
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA ADD CONTA_ID BIGINT"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_TABCONTA", "") = True Then _
            CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_UPDATE_ID_TABCONTA"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_ID_TABCONTA", "") = False Then
            SQL = "CREATE PROCEDURE SP_UPDATE_ID_TABCONTA "
            SQL = SQL & " as "
            SQL = SQL & " DECLARE @Contador AS SMALLINT"
            SQL = SQL & " SET @Contador = 0"
            SQL = SQL & " Update CONTA"
            SQL = SQL & " SET @Contador = @Contador + 1"
            SQL = SQL & " , CONTA_ID = @Contador"
            CONECTA_RETAGUARDA.Execute SQL
         End If

         CONECTA_RETAGUARDA.Execute "EXEC SP_UPDATE_ID_TABCONTA "

         Alteração_Definição_Campo_Tabela "CONTA_ID", "BIGINT NOT NULL", "CONTA", "RETAGUARDA"

         If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_CONTA", "") = False Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA ADD CONSTRAINT PK_CONTA PRIMARY KEY (CONTA_ID)"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "CONTA") = True Then _
         Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT NOT NULL", "CONTA", "RETAGUARDA"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA","BANCO", "CONTA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CONTA.BANCO'" & "," & "'BANCO_ID'" & "," & "'COLUMN'"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CONTA_BANCO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA DROP CONSTRAINT FK_CONTA_BANCO "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "BANCO_ID", "CONTA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA DROP COLUMN BANCO_ID"
      '   Alteração_Definição_Campo_Tabela "BANCO_ID", "BIGINT NOT NULL", "CONTA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_AGENCIA", "CONTA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CONTA.NUMR_AGENCIA'" & "," & "'AGENCIA_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "AGENCIA_ID", "CONTA") = True Then _
         Alteração_Definição_Campo_Tabela "AGENCIA_ID", "BIGINT NOT NULL", "CONTA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_CONTA", "CONTA") = True Then _
         Alteração_Definição_Campo_Tabela "NUMR_CONTA", "nvarchar(30) NOT NULL", "CONTA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESC_CONTA", "CONTA") = True Then _
         Alteração_Definição_Campo_Tabela "DESC_CONTA", "nvarchar(100) NOT NULL", "CONTA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_Cadastro", "CONTA") = True Then _
         Alteração_Definição_Campo_Tabela "DT_CADASTRO", "DATETIME NOT NULL", "CONTA", "RETAGUARDA"

      'If EXISTE_OBJ_BANCO("RETAGUARDA","FK_CONTA_BANCO","") = False Then
      '   SQL = "ALTER TABLE [dbo].[CONTA] WITH CHECK ADD CONSTRAINT [FK_CONTA_BANCO] FOREIGN KEY([BANCO_ID])"
      '   SQL = SQL & " References [dbo].[BANCO]([BANCO_ID])"
      '   CONECTA_RETAGUARDA.Execute SQL

      '   SQL = " ALTER TABLE [dbo].[CONTA] CHECK CONSTRAINT [FK_CONTA_BANCO]"
      '   CONECTA_RETAGUARDA.Execute SQL
      'End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CONTA_AGENCIA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CONTA] WITH CHECK ADD CONSTRAINT [FK_CONTA_AGENCIA] FOREIGN KEY([AGENCIA_ID])"
         SQL = SQL & " References [dbo].[AGENCIA]([AGENCIA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CONTA] CHECK CONSTRAINT [FK_CONTA_AGENCIA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CONTA_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CONTA] WITH CHECK ADD CONSTRAINT [FK_CONTA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CONTA] CHECK CONSTRAINT [FK_CONTA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      Else
         SQL = "CREATE TABLE [dbo].[CONTA]("
         SQL = SQL & " [CONTA_ID] [bigint] NOT NULL,"
         SQL = SQL & " [BANCO_ID] [bigint] NOT NULL,"
         SQL = SQL & " [AGENCIA_ID] [bigint] NOT NULL,"
         SQL = SQL & " [PESSOA_ID] [bigint] NOT NULL,"
         SQL = SQL & " [NUMR_CONTA] [nvarchar](30) NOT NULL,"
         SQL = SQL & " [DESC_CONTA] [nvarchar](100) NOT NULL,"
         SQL = SQL & " [DT_Cadastro] [datetime] NOT NULL,"
         SQL = SQL & " CONSTRAINT [PK_CONTA] PRIMARY KEY CLUSTERED"
         SQL = SQL & " ([CONTA_ID] Asc)"
         SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
         SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
         SQL = SQL & " ) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[CONTA] WITH CHECK ADD CONSTRAINT [FK_CONTA_BANCO] FOREIGN KEY([BANCO_ID])"
         SQL = SQL & " References [dbo].[BANCO]([BANCO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CONTA] CHECK CONSTRAINT [FK_CONTA_BANCO]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[CONTA] WITH CHECK ADD CONSTRAINT [FK_CONTA_AGENCIA] FOREIGN KEY([AGENCIA_ID])"
         SQL = SQL & " References [dbo].[AGENCIA]([AGENCIA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CONTA] CHECK CONSTRAINT [FK_CONTA_AGENCIA]"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[CONTA] WITH CHECK ADD CONSTRAINT [FK_CONTA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CONTA] CHECK CONSTRAINT [FK_CONTA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
   End If

'======================CHEQUE
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CHEQUE", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CHEQUE_ID", "CHEQUE") = True Then _
         Alteração_Definição_Campo_Tabela "CHEQUE_ID", "BIGINT NOT NULL", "CHEQUE", "RETAGUARDA"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "PK_CHEQUE", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD CONSTRAINT PK_CHEQUE PRIMARY KEY (CHEQUE_ID)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_CONTA", "CHEQUE") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CHEQUE.NUMR_CONTA'" & "," & "'CONTA_ID'" & "," & "'COLUMN'"
      
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTA_ID", "CHEQUE") = True Then _
         Alteração_Definição_Campo_Tabela "CONTA_ID", "BIGINT", "CHEQUE", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_AGENCIA", "CHEQUE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE DROP COLUMN NUMR_AGENCIA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Banco", "CHEQUE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE DROP COLUMN Banco"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Codigo_Resp", "CHEQUE") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CHEQUE.Codigo_Resp'" & "," & "'RESP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CMC7", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD CMC7 nvarchar(40) NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRAÇA", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD PRAÇA nvarchar(3) NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "RESPONSAVEL", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD RESPONSAVEL nvarchar(MAX) NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "REPASSE_ID", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD REPASSE_ID BIGINT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "REPASSE", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD REPASSE nvarchar(100) NULL"

      Alteração_Definição_Campo_Tabela "RESP_ID", "BIGINT", "CHEQUE", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD ESTABELECIMENTO_ID INT"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CHEQUE_CONTA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CHEQUE] WITH CHECK ADD CONSTRAINT [FK_CHEQUE_CONTA] FOREIGN KEY([CONTA_ID])"
         SQL = SQL & " References [dbo].[CONTA]([CONTA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CHEQUE] CHECK CONSTRAINT [FK_CHEQUE_CONTA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CHEQUE_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CHEQUE] WITH CHECK ADD CONSTRAINT [FK_CHEQUE_PESSOA] FOREIGN KEY([RESP_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CHEQUE] CHECK CONSTRAINT [FK_CHEQUE_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      Else
         SQL = "CREATE TABLE [dbo].[CHEQUE]("
         SQL = SQL & " [CHEQUE_ID] [bigint] NOT NULL,"
         SQL = SQL & " [RESP_ID] [bigint] NULL,"
         SQL = SQL & " [CONTA_ID] [bigint] NULL,"
         SQL = SQL & " [NUMR_CHEQUE] [varchar](20) NOT NULL,"
         SQL = SQL & " [SERIE_CHEQUE] [char](10) NOT NULL,"
         SQL = SQL & " [VALOR] [float] NOT NULL,"
         SQL = SQL & " [DT_EMISSAO] [datetime] NOT NULL,"
         SQL = SQL & " [DT_DEPOSITO] [datetime] NULL,"
         SQL = SQL & " [DT_COMPENSA] [datetime] NULL,"
         SQL = SQL & " [STATUS] [char](1) NULL,"
         SQL = SQL & " [NUMR_DOC] [nvarchar](30) NULL,"
         SQL = SQL & " [CMC7] [nvarchar](40) NULL,"
         SQL = SQL & " [PRAÇA] [nvarchar](3) NULL,"
         SQL = SQL & " [RESPONSAVEL] [nvarchar](MAX) NULL,"
         SQL = SQL & " [REPASSE_ID] [bigint] NULL,"
         SQL = SQL & " [REPASSE] [nvarchar](100) NULL,"
         SQL = SQL & " [ESTABELECIMENTO_ID] [INT] NULL,"
         SQL = SQL & " CONSTRAINT [PK_CHEQUE] PRIMARY KEY CLUSTERED([CHEQUE_ID] Asc"
         SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[CHEQUE] WITH CHECK ADD CONSTRAINT [FK_CHEQUE_CONTA] FOREIGN KEY([CONTA_ID])"
         SQL = SQL & " References [dbo].[CONTA]([CONTA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CHEQUE] CHECK CONSTRAINT [FK_CHEQUE_CONTA]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[CHEQUE] WITH CHECK ADD CONSTRAINT [FK_CHEQUE_PESSOA] FOREIGN KEY([RESP_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CHEQUE] CHECK CONSTRAINT [FK_CHEQUE_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
   End If

   If FSO.FileExists(App.Path & "\TXT\BANCOS.XLS") Then
      RODA_BANCO
      'Else
      '   If FSO.FileExists(App.Path & "\TXT\BANCOS.XLSX") Then
      '      RODA_BANCO
      '   End If
   End If

   If FSO.FileExists(App.Path & "\TXT\AGENCIAS.XLS") Then
      RODA_AGENCIA
      'Else
      '   If FSO.FileExists(App.Path & "\TXT\BANCOS.XLSX") Then
      '      RODA_BANCO
      '   End If
   End If

   MsgBox "ok"
End Sub

Sub RODA_BANCO()
'On Error GoTo ERRO_TRATA

   Msg = "Deseja importar cadastro de banco FENABRAM ?"
   PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      SQL3 = "BANCO,AGENCIA,CONTA,CHEQUE"

      Set oConn = New ADODB.Connection
      oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                         "FIL=excel 8.0;" & _
                         "DefaultDir=" & App.Path & "\TXT\" & ";" & _
                         "MaxBufferSize=2048;" & _
                         "PageTimeout=5;" & _
                         "DBQ=" & App.Path & "\TXT\BANCOS.XLS" & ";"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'aabre o recordset pelo nome da planilha
      TabConsulta.Open "[BancosE$]", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

      TabConsulta.MoveFirst

      If TabConsulta.EOF Then
         MsgBox "Planilha incorreta !!!"
         Exit Sub
      End If

      While Not TabConsulta.EOF

         If Not IsNull(TabConsulta.Fields(0).Value) Then
            If Trim(TabConsulta.Fields(0).Value) <> "" Then
               If IsNumeric(TabConsulta.Fields(0).Value) Then
                  If TabConsulta.Fields(0).Value > 0 Then
                     If Not IsNull(TabConsulta.Fields(1).Value) Then
                        If TabBANCO.State = 1 Then _
                           TabBANCO.Close

                        SP_PROC_BANCO TabConsulta.Fields(0).Value

                        If TabBANCO.EOF Then
                           If TabBANCO.State = 1 Then _
                              TabBANCO.Close

                           SQL = "INSERT INTO BANCO "
                              SQL = SQL & " (banco_id,codg_banco,Nome_banco)"
                           SQL = SQL & " VALUES ("
                              SQL = SQL & MAX_ID("banco_id", "BANCO", "", "", "", "")
                              SQL = SQL & ",'" & Trim(TabConsulta.Fields(0).Value) & "'"
                              SQL = SQL & ",'" & Trim(TabConsulta.Fields(1).Value) & "'"
                           SQL = SQL & ")"

                           CONECTA_RETAGUARDA.Execute SQL
                           Else
                              SQL = "update BANCO set "
                              SQL = SQL & " nome_banco = '" & Trim(TabConsulta.Fields(1).Value) & "'"
                              SQL = SQL & " where codg_banco = '" & Trim(TabConsulta.Fields(0).Value) & "'"
                              CONECTA_RETAGUARDA.Execute SQL
                        End If
                        If TabBANCO.State = 1 Then _
                           TabBANCO.Close

                        Command25.Caption = TabConsulta.Fields(0).Value
                     End If
                  End If
               End If
            End If
         End If
         DoEvents
         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      Command25.Caption = SQL3
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RODA_BANCO"
End Sub

Sub RODA_AGENCIA()
'On Error GoTo ERRO_TRATA

   Msg = "Deseja importar cadastro de agencias FENABRAM ?"
   PERGUNTA Msg, vbYesNo + 32, "Emissao NFE", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      SQL3 = "BANCO,AGENCIA,CONTA,CHEQUE"

      Set oConn = New ADODB.Connection
      oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                         "FIL=excel 8.0;" & _
                         "DefaultDir=" & App.Path & "\TXT\" & ";" & _
                         "MaxBufferSize=2048;" & _
                         "PageTimeout=5;" & _
                         "DBQ=" & App.Path & "\TXT\AGENCIAS.XLS" & ";"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'aabre o recordset pelo nome da planilha
      TabConsulta.Open "[Plan1$]", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

      TabConsulta.MoveFirst

      If TabConsulta.EOF Then
         MsgBox "Planilha incorreta !!!"
         Exit Sub
      End If

      While Not TabConsulta.EOF
         If Not IsNull(TabConsulta.Fields(0).Value) Then
            If Trim(TabConsulta.Fields(0).Value) <> "" Then
               If Not IsNull(TabConsulta.Fields(1).Value) Then
                  If Trim(TabConsulta.Fields(1).Value) <> "" Then       'codg agencia
                     If Not IsNull(TabConsulta.Fields(2).Value) Then
                        If Trim(TabConsulta.Fields(2).Value) <> "" Then 'nome agencia
                           If TabBANCO.State = 1 Then _
                              TabBANCO.Close

                           SQL = "select * from BANCO "
                           SQL = SQL & " where upper(nome_banco) = '" & Trim(TabConsulta.Fields(0).Value) & "'"
                           TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If Not TabBANCO.EOF Then
                              If TabAGENCIA.State = 1 Then _
                                 TabAGENCIA.Close
   
                              SQL = "select * from AGENCIA "
                              SQL = SQL & " where upper(numr_agencia) = '" & Trim(UCase(TabConsulta.Fields(1).Value)) & "'"
                              SQL = SQL & " and banco_id = " & TabBANCO.Fields("banco_id").Value
                              TabAGENCIA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                              If TabAGENCIA.EOF Then
                                 SQL = "INSERT INTO AGENCIA "
                                    'SQL = SQL & " (AGENCIA_ID,BANCO_ID,NUMR_AGENCIA,CODG_BANCO,NOME_AGENCIA,ENDERECO_ID)"
                                    SQL = SQL & " (AGENCIA_ID,BANCO_ID,NUMR_AGENCIA,CODG_BANCO,NOME_AGENCIA)"
                                 SQL = SQL & " VALUES ("
                                    SQL = SQL & MAX_ID("AGENCIA_ID", "AGENCIA", "", "", "", "") 'AGENCIA_ID
                                    SQL = SQL & "," & TabBANCO.Fields("banco_id").Value         'BANCO_ID
                                    SQL = SQL & ",'" & Trim(TabConsulta.Fields(1).Value) & "'"  'NUMR_AGENCIA
                                    SQL = SQL & "," & TabBANCO.Fields("codg_banco").Value       'CODG_BANCO
                                    SQL = SQL & ",'" & Trim(Replace(TabConsulta.Fields(2).Value, "'", ".")) & "'" 'NOME_AGENCIA
                                    'SQL = SQL & "," & "null"                                    'ENDERECO_ID
                                 SQL = SQL & ")"
                                 CONECTA_RETAGUARDA.Execute SQL
                              End If
                           End If
                           If TabAGENCIA.State = 1 Then _
                              TabAGENCIA.Close
   
                           Command25.Caption = TabConsulta.Fields(1).Value
                        End If
                     End If
                  End If
               End If
            End If
         End If
         DoEvents
         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      Command25.Caption = SQL3
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RODA_AGENCIA"
End Sub

Private Sub SSPanel1_Click()
'On Error GoTo ERRO_TRATA

   Dim Str_Database As String
   Dim Str_Data As String
   Dim Str_Use As String
   Dim Str_Master As String
   Dim Str_Detail As String

   Dim ConData As New ADODB.Connection
   Const App_Name = "Criador de Banco de dados SQL Server"

   SqL2 = InputBox("Informe Nome Servidor.", "SHF INFORMÁTICA")

   If SqL2 = vbNullString Then
      MsgBox "Informe o nome do Servidor.", vbInformation, ap_name
      Exit Sub
   End If

'Este demo funciona para maquinas locais
'Voce pode modificar o demo para usar o SQL Server Remoto alterando a string de conexão

   ConData.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist " & _
                "Security Info=False;Initial Catalog=master;Data Source=.\SQLEXPRESS"

   SqL2 = ".\SQLEXPRESS" 'servidor local sql server

   Str_Database = InputBox("Informe Nome Banco de Dados a ser criado.", "SHF INFORMÁTICA")

   If Str_Database = vbNullString Then
      MsgBox "Informe o nome do banco de dados.", vbInformation, App_Name
      Exit Sub
   End If

   Str_Data = "Create database " & Str_Database
   ConData.Execute (Str_Data)

   If ConData.State = 1 Then _
      ConData.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSPanel1_Click"
End Sub

Sub RODA_IMPORTA_NCM()
'On Error GoTo ERRO_TRATA

   CONT_N = 0

   Msg = "Deseja planilha NCM ?"
   PERGUNTA Msg, vbYesNo + 32, "IMPORTAÇÃO NCM", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      SQL3 = "LEI DO IMPOSTO 12.741"

      Set oConn = New ADODB.Connection
      oConn.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                         "FIL=excel 8.0;" & _
                         "DefaultDir=" & App.Path & "\TXT\" & ";" & _
                         "MaxBufferSize=2048;" & _
                         "PageTimeout=5;" & _
                         "DBQ=" & App.Path & "\TXT\NCMimporta.XLS" & ";"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'aabre o recordset pelo nome da planilha
      TabConsulta.Open "[Plan1$]", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

      TabConsulta.MoveFirst

      If TabConsulta.EOF Then
         MsgBox "Planilha incorreta !!!"
         Exit Sub
      End If

      While Not TabConsulta.EOF

         If Not IsNull(TabConsulta.Fields(0).Value) Then
            If Trim(TabConsulta.Fields(0).Value) <> "" Then
               If IsNumeric(TabConsulta.Fields(0).Value) Then
                  If TabConsulta.Fields(0).Value > 0 Then
                     If Not IsNull(TabConsulta.Fields(1).Value) Then

                        SqL2 = "" & Replace(TabConsulta.Fields(0).Value, ".", "")
                        PERC_N = 0 & TabConsulta.Fields(2).Value
                        NOME_A = "" & Replace(TabConsulta.Fields(1).Value, "'", "´")
                        NOME_A = "" & Replace(NOME_A, ",", ";")

                        If TabTemp.State = 1 Then _
                           TabTemp.Close

                        SQL = "select * from TABNCM "
                        SQL = SQL & " where codg_ncm = '" & Trim(SqL2) & "'"
                        TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If TabTemp.EOF Then
                           If TabTemp.State = 1 Then _
                              TabTemp.Close

                           SQL = "INSERT INTO TABNCM "
                           SQL = SQL & " VALUES ("
                              SQL = SQL & "'" & Trim(SqL2) & "'"
                              SQL = SQL & ",'" & Trim(NOME_A) & "'"
                              SQL = SQL & "," & tpMOEDA(PERC_N)
                           SQL = SQL & ")"

                           CONECTA_RETAGUARDA.Execute SQL

                           CONT_N = CONT_N + 1

                           Else
                              SQL = "update TABNCM set "
                              SQL = SQL & " aliquota_ncm = " & tpMOEDA(PERC_N)
                              SQL = SQL & " where codg_ncm = '" & Trim(SqL2) & "'"
                              CONECTA_RETAGUARDA.Execute SQL
                        End If
                        If TabTemp.State = 1 Then _
                           TabTemp.Close

                        Command12.Caption = CONT_N
                     End If
                  End If
               End If
            End If
         End If
         DoEvents
         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      Command12.Caption = SQL3
   End If
   SQL3 = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "RODA_BANCO"
End Sub
Private Sub CMDSU_Click()

   Proc_n = 0
   Novos_n = 0
   At_n = 0
   
   'Abrir o arquivo do Excel
   Set xlw = xl.Workbooks.Open("C:\MEGASIM\TXT\IMP.XLSX")

   ' definir qual a planilha de trabalho
   xlw.Sheets("Plan1").Select

   DESCRICAO = "a"
   Proc_n = 0

   While Trim(DESCRICAO) <> ""
      Proc_n = Proc_n + 1
      CMDSU.Caption = "Processados = " & Proc_n
      DoEvents

      DESCRICAO = xlw.Application.Cells(Proc_n, 1).Value

      If Trim(DESCRICAO) <> "" Then
         CNPJCPF = Replace(Trim(xlw.Application.Cells(Proc_n, 2).Value), "'", "")
         CNPJCPF = Replace(CNPJCPF, ".", "")
         CNPJCPF = Replace(CNPJCPF, "-", "")
         CNPJCPF = Replace(CNPJCPF, "/", "")

         spPessoa 1, 0, CNPJCPF, DESCRICAO, DESCRICAO, "A"

         CHECA_CLIENTE
      End If
   Wend

   ' Fechar a planilha sem salvar alterações
   ' Para salvar mude False para True
   xlw.Close False

   ' Liberamos a memória
   Set xlw = Nothing
   Set xl = Nothing

   MsgBox "Ok"

End Sub

Sub CHECA_CLIENTE()
'On Error GoTo ERRO_TRATA

   If Trim(CNPJCPF) <> "" Then
      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select cliente_id,PESSOA_id from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF) & "'"
      SQL = SQL & " and status = 'A'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         CLIENTE_ID_N = TabCliente.Fields("cliente_id").Value
         PESSOA_ID_N = TabCliente.Fields("PESSOA_id").Value
         Else
            If TabCliente.State = 1 Then _
               TabCliente.Close

            '============================
            PESSOA_ID_N = MAX_ID("pessoa_id", "pessoa", "", "", "", "")
            spPessoa 1, 0, Trim(CNPJCPF), Trim(DESCRICAO), Trim(DESCRICAO), "A"

            CLIENTE_ID_N = MAX_ID("cliente_id", "cliente", "", "", "", "")

            SQL = "insert into CLIENTE "
               SQL = SQL & "(cliente_id,pessoa_id,empresa_id,vendedor_id,"
               SQL = SQL & " cgccpf,nome,razao_social,dt_cad,status,ie,estrangeiro)"
            SQL = SQL & " values("
               SQL = SQL & CLIENTE_ID_N                        'cliente_id
               SQL = SQL & "," & PESSOA_ID_N                   'pessoa_id
               SQL = SQL & "," & EMPRESA_ID_N                  'empresa_id
               SQL = SQL & "," & VENDEDOR_ID_N                 'vendedor_id
               SQL = SQL & ",'" & Trim(CNPJCPF) & "'"   'cgccpf
               SQL = SQL & ",'" & Trim(DESCRICAO) & "'"     'nome
               SQL = SQL & ",'" & Trim(DESCRICAO) & "'"     'razao_social
               SQL = SQL & ",'" & Now & "'"              'dt_cad
               SQL = SQL & ",'A'"                              'status
               SQL = SQL & ",'ISENTO'"                         'ie)
               SQL = SQL & ",0"                                'estrangeiro
            SQL = SQL & " )"
            CONECTA_RETAGUARDA.Execute SQL
            '============================
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      If TabUSU.State = 1 Then _
         TabUSU.Close

      SQL = "select nome,pessoa_id from USUARIO"
      SQL = SQL & " where cpf = '" & Trim(CNPJCPF) & "'"
      TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabUSU.EOF Then
         SQL = "INSERT INTO USUARIO "
            SQL = SQL & " (empresa_id,usuario_id,Nome,Cpf,Status,Pessoa_id,FUNCIONARIO) "
         SQL = SQL & " VALUES ("
            SQL = SQL & EMPRESA_ID_N
            SQL = SQL & "," & MAX_ID("usuario_id", "usuario", "empresa_id", "1", "", "")
            SQL = SQL & ",'" & Trim(DESCRICAO) & "'"
            SQL = SQL & ",'" & Trim(CNPJCPF) & "'"
            SQL = SQL & ",'TRUE'"
            SQL = SQL & "," & PESSOA_ID_N
            SQL = SQL & ",'true'"
            SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabUSU.State = 1 Then _
         TabUSU.Close
   End If

   PESSOA_ID_N = 0
   CLIENTE_ID_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CHECA_CLIENTE"
End Sub

Private Sub EEE()

   Dim PRECO_CUSTO_N  As Double

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select PEDIDO.tabelapreco_id, PEDIDO.DT_REQ, PEDIDOITEM.PRODUTO_ID, "
   SQL = SQL & " PEDIDOITEM.PRECO_CUSTO, ITEMLANCAMENTO.FORMAPAGTO_ID,PEDIDOITEM.seq_id,PEDIDOITEM.pedido_id"
   SQL = SQL & " from PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " INNER JOIN ITEMLANCAMENTO "
   SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = ITEMLANCAMENTO.NUMR_DOC"

   SQL = SQL & " where dt_req >= '01/05/2016 00:00:00'"
'--and   dt_req <= '31/05/2016 23:00:00'

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF

      PRECO_CUSTO_N = 0 & TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(TabTemp.Fields("produto_id").Value, TabTemp.Fields("tabelapreco_id").Value, TabTemp.Fields("FORMAPAGTO_ID").Value)
Command1.Caption = TabTemp.Fields("pedido_id").Value
DoEvents
      If TabTemp.Fields("preco_custo").Value < PRECO_CUSTO_N Then
'      MsgBox TabTemp.Fields("dt_req").Value
'MsgBox "Produto = " & TabTemp.Fields("produto_id").Value & "   preço custo tabela =  " & _
TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(TabTemp.Fields("produto_id").Value, _
TabTemp.Fields("tabelapreco_id").Value, TabTemp.Fields("FORMAPAGTO_ID").Value) & "  preço custo no pedidoitem = " & TabTemp.Fields("preco_custo").Value & "   data = " & TabTemp.Fields("dt_req").Value

         SQL = "update PEDIDOITEM set preco_custo = " & tpMOEDA(PRECO_CUSTO_N)
         SQL = SQL & " where pedido_id = " & TabTemp.Fields("pedido_id").Value
         SQL = SQL & " and seq_id = " & TabTemp.Fields("seq_id").Value
         CONECTA_RETAGUARDA.Execute SQL
      
'MsgBox SQL
      End If
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

MsgBox "ok"
End Sub




Private Sub Command5_Click()
   
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
   SQL = "select * from import"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      NUMR_ID_N = MAX_ID("produto_id", "produto", "", "", "", "")
      SQL = "insert into PRODUTO "
         SQL = SQL & "("
            SQL = SQL & "PRODUTO_ID,EMPRESA_ID,CODG_PRODUTO,FORNECEDOR_ID,DESCRICAO"
            SQL = SQL & ",REFERENCIA,FAMILIAPRODUTO_ID,UNIDADE_MEDIDA,CODG_BARRA,SITUACAO"
            SQL = SQL & ",SITUACAO_TRIBUTARIA,ALIQUOTA_ICMS,PERC_DESCONTO,TIPO_PROD,CODG_NCM"
            SQL = SQL & ",COMP_TRIBUTARIA,PRECO_CUSTO_ANTERIOR,qtd_ped_anterior,PRECO_CUSTO"
            SQL = SQL & ",PRECO_ATACADO,PRECO_Venda,PERCIVA,DT_CADASTRO,PERC_COMIS,PATH_IMAGEM"
            SQL = SQL & ",ORIGEM_MERCADO,LOCACAO,PRECO_VAREJO_ANTERIOR,PRECO_ATACADO_ANTERIOR"
            SQL = SQL & ",EMBALAGEM,USUARIO_ID,QTD_MINIMO,QTD_MAXIMO,DT_ULT_VENDA,DT_ULT_COMPRA"
            SQL = SQL & ",PESO_LIQUIDO,PESO_BRUTO,TAMANHO,MARCA_ID,PRODUTO_BALANCA,PERMITE_DESCONTO"
            SQL = SQL & ",CONCEDER_PRODUCAO,PERC_COMPOE_VENDA"
         SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & NUMR_ID_N             'produto_id
         SQL = SQL & "," & EMPRESA_ID_N                                          '[EMPRESA_ID]
         SQL = SQL & ",'" & Trim(NUMR_ID_N) & "'"   'CODG_PRODUTO
         SQL = SQL & ",0"                                                        'FORNECEDOR_ID]
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("nome").Value) & "'"         '[DESCRICAO]
         SQL = SQL & ",'" & Trim("") & "'"                                       '[REFERENCIA]
         SQL = SQL & ",1"                                                       'FAMILIAPRODUTO_ID]
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("un").Value) & "'"           'UNIDADE_MEDIDA]
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("Codbarra").Value) & "'"     'CODG_BARRA]
         SQL = SQL & ",'A'"                                                      'SITUACAO]
         SQL = SQL & ",'00'"                                                     'SITUACAO_TRIBUTARIA]
         SQL = SQL & ",17"                                                       'ALIQUOTA_ICMS]
         SQL = SQL & ",0"                                                        'PERC_DESCONTO]
         SQL = SQL & ",1"                                                        'TIPO_PROD]
         SQL = SQL & ",'" & Trim(TabConsulta.Fields("ncm").Value) & "'"          'CODG_NCM]
         SQL = SQL & ",0"                                                       'COMP_TRIBUTARIA]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PRECO_CUSTO_ANTERIOR]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'qtd_ped_anterior]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PRECO_CUSTO]
         SQL = SQL & ",'" & tpMOEDA(TabConsulta.Fields("valvenda").Value) & "'"  'PRECO_ATACADO]
         SQL = SQL & ",'" & tpMOEDA(TabConsulta.Fields("valvenda").Value) & "'"  'PRECO_Venda]
         SQL = SQL & ",'0'"                                                      'PERCIVA]
         SQL = SQL & ",'" & Now & "'"                                            'DT_CADASTRO]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PERC_COMIS]
         SQL = SQL & ",'" & tpMOEDA("") & "'"                                    'PATH_IMAGEM]
         SQL = SQL & ",'" & Trim(0) & "'"                                        'ORIGEM_MERCADO]
         SQL = SQL & ",'" & Trim("") & "'"                                       'LOCACAO]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PRECO_VAREJO_ANTERIOR]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PRECO_ATACADO_ANTERIOR]
         SQL = SQL & ",'" & Trim("") & "'"                                         'EMBALAGEM]
         SQL = SQL & ",'144'"                                                    'USUARIO_ID]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'QTD_MINIMO]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'QTD_MAXIMO]
         SQL = SQL & ",NULL"                                                   'DT_ULT_VENDA]
         SQL = SQL & ",NULL"                                                   'DT_ULT_COMPRA]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PESO_LIQUIDO]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PESO_BRUTO]
         SQL = SQL & ",'" & Trim("") & "'"                                       'TAMANHO]
         SQL = SQL & ",0"                                                        'MARCA_ID]
         SQL = SQL & ",0"                                                        'PRODUTO_BALANCA]
         SQL = SQL & ",0"                                                        'PERMITE_DESCONTO]
         SQL = SQL & ",0"                                                        'CONCEDER_PRODUCAO]
         SQL = SQL & ",'" & tpMOEDA(0) & "'"                                     'PERC_COMPOE_VENDA]
      SQL = SQL & ")"

      CONECTA_RETAGUARDA.Execute SQL
      TabConsulta.MoveNext
   Wend

End

   If TabConsulta.State = 1 Then _
      TabConsulta.Close
   CONT_N = 0
   CRITERIO_A = ""
   SQL = "select cliente_id,ie from CLIENTE "
   SQL = SQL & " where ie is not null"
   SQL = SQL & " and ie <> 'ISENTO'"
   SQL = SQL & " ORDER BY PESSOA_ID"
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF

      If Trim(TabConsulta.Fields("ie").Value) <> "" Then
         CRITERIO_A = Trim(Replace(TabConsulta.Fields("ie").Value, "-", ""))
         CRITERIO_A = Trim(Replace(CRITERIO_A, ".", ""))
         CRITERIO_A = Trim(Replace(CRITERIO_A, ",", ""))
         CRITERIO_A = Trim(Replace(CRITERIO_A, ";", ""))

'MsgBox TabConsulta.Fields("ie").Value & "    " & CRITERIO_A

SQL = "update cliente set "
SQL = SQL & " ie = '" & Trim(CRITERIO_A) & "'"
SQL = SQL & " where cliente_id = " & TabConsulta.Fields("cliente_id").Value
CONECTA_RETAGUARDA.Execute SQL

      End If

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
MsgBox "ok"
End Sub



