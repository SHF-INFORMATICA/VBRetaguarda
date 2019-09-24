VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form frmImporta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de Arquivos"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmImporta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Importação Geral"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.PictureBox DLG 
         Height          =   395
         Left            =   6360
         Picture         =   "frmImporta.frx":5C12
         ScaleHeight     =   388.235
         ScaleMode       =   0  'User
         ScaleWidth      =   366.667
         TabIndex        =   10
         Top             =   680
         Width           =   395
      End
      Begin VB.TextBox txtRetorno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1800
         TabIndex        =   1
         Top             =   680
         Width           =   4815
      End
      Begin Threed.SSCommand cmdGerar 
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Gera Imp. IBGE"
      End
      Begin Threed.SSCommand cmdSair 
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sair"
      End
      Begin Threed.SSCommand cmdCep 
         Height          =   495
         Left            =   4920
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Gera Imp. CEP"
      End
      Begin VB.Label Label2 
         Caption         =   "Caminho Arquivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblcep 
         Caption         =   "Registro CEP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblibge 
         Caption         =   "Registro IBGE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblcontcep 
         Caption         =   "contcep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label lblcontibge 
         Caption         =   "contibge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   2400
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cont As Integer
Dim strUF As String
Dim strCidade As String
Dim strIBGE As String
Dim strCEP As String
Dim strRegistro As String

Private Sub cmdGerar_Click()
    'Importando Codigo IBGE para Sistema EJS
    'primeiro tem que executar a importacao do cep depois do codigo ibge
    'cmdCep_Click
    Cont = 0
    Open txtRetorno.Text For Input As #1
    Do While Not EOF(1)
       DoEvents
       Line Input #1, strRegistro
       strUF = Mid$(strRegistro, 1, 2) 'UF
       strIBGE = Mid$(strRegistro, 4, 7) 'Codigo IBGE
       strCidade = Mid$(strRegistro, 12, 70) 'Cidade
       TabCEP.Open "select * from CEP Where uf = '" & strUF & "' and Cidade = '" & TiraAcento(strCidade) & "'", CONECTA_RETAGUARDA, , , adCmdText
       If Not TabCEP.EOF Then
          CONECTA_RETAGUARDA.Execute "UPDATE CEP SET IBGE_ID = " & numeros(strIBGE) & " where uf = '" & strUF & "' and Cidade = '" & TiraAcento(strCidade) & "'"
          Cont = Cont + 1
       End If
       TabCEP.Close
       lblcontibge.Caption = Cont
    Loop
    Close #1    ' Close file.
    MsgBox "Foram Importador" & Cont & " Registro de IBGE"
    lblcontibge.Caption = ""
    Exit Sub
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdCep_Click()
    'Importando CEP
    CONECTA_RETAGUARDA.Execute "Delete from CEP"
    Open txtRetorno.Text For Input As #1
    Do While Not EOF(1)
       DoEvents
       Line Input #1, strRegistro
       Cont = Cont + 1
       strUF = Mid$(strRegistro, 1, 2) 'UF
       strCEP = Mid$(strRegistro, 4, 8) 'Codigo CEP
       strCidade = Mid$(strRegistro, 13, 70) 'Cidade
       CONECTA_RETAGUARDA.Execute "INSERT INTO CEP (CEP_id, uf, cidade) VALUES (" & numeros(strCEP) & ",'" & strUF & "','" & TiraAcento(strCidade) & "')"
       lblcontcep.Caption = Cont
    Loop
    Close #1    ' Close file.
    MsgBox "Foram Importador" & Cont & " Registro de Ceps"
    lblcontcep.Caption = ""
    Exit Sub
    
    
End Sub

Private Sub DLG_Click()
   frmINICIO.Dialogo.DialogTitle = "Selecionar Caminho Arquivo!"
   frmINICIO.Dialogo.Filter = "*.txt;*.xls"
   frmINICIO.Dialogo.ShowOpen
   If frmINICIO.Dialogo.FileName <> "" Then
      txtRetorno.Text = frmINICIO.Dialogo.FileName
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   'If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub


