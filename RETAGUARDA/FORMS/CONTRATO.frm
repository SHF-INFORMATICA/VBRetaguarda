VERSION 5.00
Begin VB.Form frmCONTRATO 
   BackColor       =   &H00800000&
   Caption         =   "Controle de Documentos"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CONTRATO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   12690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   1335
      Left            =   10200
      Picture         =   "CONTRATO.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdRel 
      Caption         =   "&Relatórios"
      Height          =   1335
      Left            =   7680
      Picture         =   "CONTRATO.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdFinanceiro 
      Caption         =   "&Financeiro"
      Height          =   1335
      Left            =   5160
      Picture         =   "CONTRATO.frx":2C2E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdContrato 
      Caption         =   "C&ontratos"
      Height          =   1335
      Left            =   2640
      Picture         =   "CONTRATO.frx":3F2E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdCadastro 
      Caption         =   "&Cadastros"
      Height          =   1335
      Left            =   120
      Picture         =   "CONTRATO.frx":53CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmCONTRATO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCadastro_Click()
   CRITERIO = "cadastro"
   frmCONTRATOOPCAO.Show 1
End Sub

Private Sub cmdContrato_Click()
   CRITERIO = "contrato"
   frmCONTRATOOPCAO.Show 1
End Sub

Private Sub cmdFinanceiro_Click()
   CRITERIO = "financeiro"
   frmCONTRATOOPCAO.Show 1
End Sub

Private Sub cmdRel_Click()
   CRITERIO = "relatorio"
   frmCONTRATOOPCAO.Show 1
End Sub

Private Sub cmdSair_Click()
   End
End Sub
