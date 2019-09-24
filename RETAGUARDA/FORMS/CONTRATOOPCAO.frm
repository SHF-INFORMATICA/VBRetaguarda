VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCONTRATOOPCAO 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "CONTRATOOPCAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LISTA 
      Height          =   3495
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6165
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   17639
      EndProperty
   End
End
Attribute VB_Name = "frmCONTRATOOPCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   If CRITERIO = "cadastro" Then
      Set Item = LISTA.ListItems.Add(, "seq." & 1, "Cadastro Cliente")
      Set Item = LISTA.ListItems.Add(, "seq." & 2, "Cadastro Fornecedor")
   End If

   If CRITERIO = "contrato" Then
      Set Item = LISTA.ListItems.Add(, "seq." & 1, "Gerar Contrato")
      Set Item = LISTA.ListItems.Add(, "seq." & 2, "Consultar Contratos")
   End If

   If CRITERIO = "financeiro" Then
      Set Item = LISTA.ListItems.Add(, "seq." & 1, "Contas a Pagar")
      Set Item = LISTA.ListItems.Add(, "seq." & 2, "Contas a Receber")
      Set Item = LISTA.ListItems.Add(, "seq." & 3, "Controle Cheque")
   End If

   If CRITERIO = "relatorio" Then
      Set Item = LISTA.ListItems.Add(, "seq." & 1, "Recibo Autorização para Escritura")
      'Set Item = LISTA.ListItems.Add(, "seq." & 2, "Consultar Contratos")
   End If

End Sub

Private Sub LISTA_DblClick()
   ABRE_BANCO_MEGASIM NOME_BANCO_DADOS

   If CRITERIO = "cadastro" Then
      If LISTA.SelectedItem.Index = 1 Then _
         frmCADASTROCLIENTE.Show 1

      If LISTA.SelectedItem.Index = 2 Then _
         frmCADASTROFORNECEDOR.Show 1
   End If

   If CRITERIO = "financeiro" Then
      If LISTA.SelectedItem.Index = 1 Then
         SINAL_INDICADOR_N = 2
         frmFINGERALANC.Show 1
      End If
      If LISTA.SelectedItem.Index = 2 Then
         SINAL_INDICADOR_N = 1
         frmFINGERALANC.Show 1
      End If
      If LISTA.SelectedItem.Index = 2 Then _
         frmCHEQUECONSULTA.Show 1
   End If

   If CONECTA_RETAGUARDA.State = 1 Then _
      CONECTA_RETAGUARDA.Close
End Sub
