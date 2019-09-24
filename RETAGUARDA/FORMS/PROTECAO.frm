VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.Form frmPROTECAO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gerar Chave"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PROTECAO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChave 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5775
   End
   Begin MSMask.MaskEdBox txtData 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Código Verificador"
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmPROTECAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGUID As GUID) As Long
   Private Declare Function StringfromGUID2 Lib "OLE32.DLL" (pGUID As GUID, _
   ByVal PointerToString As Long, ByVal MaxLength As Long) As Long
   Const GUI = "7FABFB3B-AF2F-42F5-9655-2846FC0E0D8C"

Private Type GUID
  Guid1 As Long
  Guid2 As Long
  Guid3 As Long
  Guid4(0 To 7) As Byte
End Type

Public Function CreateGUID() As String
   Dim udtGUID As GUID
   Dim sGUID As String
   Dim lResult As Long

   lResult = CoCreateGuid(udtGUID)

   If lResult Then
      sGUID = ""
      Else
         sGUID = String$(38, 0)
         StringfromGUID2 udtGUID, StrPtr(sGUID), 39
   End If

   CreateGUID = sGUID
End Function

Private Sub Form_Load()
   txtData.PromptInclude = False
   txtData.Text = DMA("10/" & Month(Date) + 1 & "/" & Year(Date))
   txtData.PromptInclude = True
End Sub

Private Sub cmdGerar_Click()
   txtChave.Text = CreateGUID
End Sub

Private Sub cmdRegistrar_Click()
   If Trim(txtChave.Text) = "" Then
      MsgBox "Informe a chave de registro."
      Else: Main
   End If
End Sub

Sub Main()
   txtData.PromptInclude = True

MsgBox Command

   If Command = "Ativar" Then
      AtivaAplicativo (CreateGUID)
      Else: DesativarEm txtData.Text
   End If

   End
End Sub

Public Sub DesativarEm(Data As Date)

   'Desativa o programa na data informada
   Dim chave As String

   'Gera a chave com base no codigo identificador do usuario
   chave = Left(GUI, 8)
   'chave = Left(txtChave.Text, 8)

   'Se a chave for invalida encerra a aplicação
   If GetSetting("MEGASIM", "Security", chave, GUI) <> GUI Then
      MsgBox "Não é possivel executar a aplicação entre em contado com o suporte técnico", vbCritical, _
      "Erro de Validação de chave : A-1"
      End
   End If

   'Se a data expirar, desativa o aplicativo
   If Date >= Data Then
      'grava um valor invalido na chave do registro
      SaveSetting "MEGASIM", "Security", chave, "A-1"

      MsgBox "O periodo de demonstração terminou ! " & vbCrLf & _
      " Para adquirir o sistema entre em contato com seu revendedor", vbCritical, "Erro Interno"
      End
   End If

End Sub

Public Sub AtivaAplicativo(codigo As String)

   Dim chave As String

   chave = Left(codigo, 8)

   If Command = "Libera" Then _
      SaveSetting "MEGASIM", "Security", chave, codigo

End Sub
