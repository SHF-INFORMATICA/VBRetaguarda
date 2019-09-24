VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmAtGlobal 
   Caption         =   "Rotinas do PUTO"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AtGlobal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand SSCommand1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ActiveColors    =   -1  'True
      Caption         =   "Campos na tabela MFA010 [MFAINDFINAL] [int] NULL,"
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   735
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ActiveColors    =   -1  'True
      Caption         =   "Campos (indPres,idDest) tabela NF"
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      _Version        =   262144
      CaptionStyle    =   1
      ActiveColors    =   -1  'True
      Caption         =   "Atualiza sp banco Global"
   End
End
Attribute VB_Name = "frmAtGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand1_Click()
'On Error GoTo ERRO_TRATA

   SQL = "1 - Criar os seguintes Campos na tabela MFA010 do banco0 de dados Global :"
   SQL = SQL & " [MFAINDFINAL] [int] NULL,"
   SQL = SQL & " [MFAIDDEST] [int] NULL,"
   SQL = SQL & " [MFAINDPRES] [int] NULL,"
   SQL = SQL & " [MFACHAVEREFNFE] [varchar](100) NULL"
   SQL = SQL & " 2 - Alterar a Seguinte Procedure spNFeCabeNota executando em todos os clientes :"
   SQL = SQL & " obs.: Fazer backup antes nos bancos e testar apos criar em um cliente para ver se nao da problema, mesmo ainda ele nao pegando essas informações no xml,"
   SQL = SQL & " pois quando ele importar os valores serão nulos"

   SQL = SQL & " MFAFINNFE, int null  = Finalidade da Emissao"

SQL = SQL & " 1=NF-e normal;"
SQL = SQL & " 2=NF-e complementar;"
SQL = SQL & " 3=NF-e de ajuste;"
SQL = SQL & " 4=Devolução de mercadoria."

MsgBox SQL

   Msg = "Confirma?"
   PERGUNTA Msg, vbYesNo + 32, "Cadastro Cliente NFE", "DEMO.HLP", 1000
   If RESPOSTA = vbYes Then
      If CONECTA_GLOBAL.State = 1 Then _
         CONECTA_GLOBAL.Close

      ABRE_BANCO_GLOBAL

      If CONECTA_GLOBAL.State <> 1 Then
         MsgBox "Banco GLOBAL não conectado."
         Exit Sub
      End If

      If EXISTE_OBJ_BANCO("GLOBAL", "MFA010", "") = True Then
         If EXISTE_CAMPO_TABELA("GLOBAL", "MFAINDFINAL", "MFA010") = False Then _
            CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD MFAINDFINAL INT"

         If EXISTE_CAMPO_TABELA("GLOBAL", "MFAIDDEST", "MFA010") = False Then _
            CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD MFAIDDEST INT"

         If EXISTE_CAMPO_TABELA("GLOBAL", "MFAINDPRES", "MFA010") = False Then _
            CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD MFAINDPRES INT"

         If EXISTE_CAMPO_TABELA("GLOBAL", "MFACHAVEREFNFE", "MFA010") = False Then _
            CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD MFACHAVEREFNFE VARCHAR(100)"

         If EXISTE_CAMPO_TABELA("GLOBAL", "MFAFINNFE", "MFA010") = False Then _
            CONECTA_GLOBAL.Execute "ALTER TABLE MFA010 ADD MFAFINNFE INT"
      End If
      If CONECTA_GLOBAL.State = 1 Then _
         CONECTA_GLOBAL.Close
      MsgBox "OK, CRIADO OS CAMPOS."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSCommand1_Click"
End Sub

Private Sub SSCommand2_Click()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "NF", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "indPres", "NF") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD indPres INT"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "idDest", "NF") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE NF ADD idDest INT"

      MsgBox "OK, CRIADO OS CAMPOS, PUTOOOOOOOOOOOOOO"
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SSCommand2_Click"
End Sub

