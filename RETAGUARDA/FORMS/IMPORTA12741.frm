VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmIMPORTA12741 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação IBPTax"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IMPORTA12741.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBusca 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   450
   End
   Begin VB.TextBox txtPath 
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
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5295
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   1270
      ButtonWidth     =   2461
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
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
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar"
            Key             =   "gravar"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":9EFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":B051
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":B4A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "IMPORTA12741.frx":D31A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblAt 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   3345
      TabIndex        =   6
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblConta 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   270
      TabIndex        =   5
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Arquivo IBPTax"
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
      TabIndex        =   4
      Top             =   720
      Width           =   6090
   End
   Begin VB.Label lbl_caminho_xml_demo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Arquivo Importar:"
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
      Left            =   180
      TabIndex        =   3
      Top             =   1365
      Width           =   2010
   End
End
Attribute VB_Name = "frmIMPORTA12741"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   LIMPA_TELA
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TELA
      Case "gravar"
         IMPORTA_IBPTax
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub cmdBusca_Click()
'On Error GoTo ERRO_TRATA

   frmINICIO.Dialogo.FileName = ""
   frmINICIO.Dialogo.InitDir = App.Path
   frmINICIO.Dialogo.DialogTitle = "Importação arquivo"
   frmINICIO.Dialogo.Filter = "*.csv;*.txt"
   frmINICIO.Dialogo.ShowOpen
   If frmINICIO.Dialogo.FileName <> "" Then _
      txtPath.Text = frmINICIO.Dialogo.FileName

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdBusca_Click"
End Sub

Sub IMPORTA_IBPTax()
'On Error GoTo ERRO_TRATA

   If frmINICIO.Dialogo.FileName = "" Then _
      Exit Sub

   Dim Arq_Conteudo(0 To 7)   As String
   Dim sLine                  As String
   Dim VALOR_ALIQNAC          As Double
   Dim VALOR_ALIQIMP          As Double

'=====================================================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "IBPTax", "") = False Then

      SQL = "CREATE TABLE [dbo].[IBPTax]("
      SQL = SQL & " [CODG_NCM] [nvarchar](10) NOT NULL,"
      SQL = SQL & " [EX_TARIFARIO] [nvarchar](2) NOT NULL,"
      SQL = SQL & " [TABELA] [INT] NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](MAX) NOT NULL,"
      SQL = SQL & " [ALIQNAC] [FLOAT] NOT NULL,"
      SQL = SQL & " [ALIQIMP] [FLOAT] NOT NULL"

      'SQL = SQL & "  CONSTRAINT [PK_IBPTax] PRIMARY KEY CLUSTERED([CODG_NCM] Asc"
      'SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      'SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

   End If
'=====================================================
   CONT_N = 0
   NUMR_SEQ_N = 0
   ARQ_TXT = FreeFile
   Open txtPath.Text For Input As ARQ_TXT
   Do While Not EOF(ARQ_TXT)
      DoEvents

      Line Input #ARQ_TXT, sLine
      ParseToArray sLine, Arq_Conteudo()

      If IsNumeric(Arq_Conteudo(0)) Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

CRITERIO_A = Replace(Arq_Conteudo(3), "'", ".")
CRITERIO_A = Replace(CRITERIO_A, ",", ";")

VALOR_ALIQNAC = Replace(Arq_Conteudo(4), ".", ",")
VALOR_ALIQIMP = Replace(Arq_Conteudo(5), ".", ",")

         SQL = "select * from IBPTAX WITH (NOLOCK)"
         SQL = SQL & " where codg_ncm = '" & Trim(Arq_Conteudo(0)) & "'"
         SQL = SQL & " and EX_TARIFARIO = '" & Trim(Arq_Conteudo(1)) & "'"
         SQL = SQL & " and tabela = '" & Trim(Arq_Conteudo(2)) & "'"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabTemp.EOF Then
            NUMR_SEQ_N = NUMR_SEQ_N + 1
            SQL = "update IBPTax set "
               SQL = SQL & "EX_TARIFARIO = '" & Trim(Arq_Conteudo(1)) & "'"   'EX_TARIFARIO
               SQL = SQL & ",TABELA = " & Trim(Arq_Conteudo(2))               'TABELA
               SQL = SQL & ",DESCRICAO= '" & Trim(CRITERIO_A) & "'"           'DESCRICAO
               SQL = SQL & ",ALIQNAC = " & tpMOEDA(VALOR_ALIQNAC)             'ALIQNAC
               SQL = SQL & ",ALIQIMP = " & tpMOEDA(VALOR_ALIQIMP)             'ALIQIMP
            SQL = SQL & " where codg_ncm = '" & Trim(Arq_Conteudo(0)) & "'"
            SQL = SQL & " and EX_TARIFARIO = '" & Trim(Arq_Conteudo(1)) & "'"
            SQL = SQL & " and tabela = '" & Trim(Arq_Conteudo(2)) & "'"
            lblAt.Caption = "Atualizados = " & NUMR_SEQ_N
            Else
               CONT_N = CONT_N + 1
               SQL = "insert into IBPTax values( "
                  SQL = SQL & "'" & Trim(Arq_Conteudo(0)) & "'"   'CODG_NCM
                  SQL = SQL & ",'" & Trim(Arq_Conteudo(1)) & "'"  'EX_TARIFARIO
                  SQL = SQL & "," & Trim(Arq_Conteudo(2))         'TABELA
                  SQL = SQL & ",'" & Trim(CRITERIO_A) & "'"       'DESCRICAO
                  SQL = SQL & "," & tpMOEDA(VALOR_ALIQNAC)        'ALIQNAC
                  SQL = SQL & "," & tpMOEDA(VALOR_ALIQIMP)        'ALIQIMP
               SQL = SQL & ")"
               lblConta.Caption = "Incluídos = " & CONT_N
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CONECTA_RETAGUARDA.Execute SQL
      End If
   Loop
   Close #1

   MsgBox "Processo realizado com sucesso."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "IMPORTA_IBPTax"
End Sub

Sub LIMPA_TELA()
   txtPath.Text = ""
   lblConta.Caption = ""
   lblAt.Caption = ""
   CONT_N = 0
   NUMR_SEQ_N = 0
End Sub
