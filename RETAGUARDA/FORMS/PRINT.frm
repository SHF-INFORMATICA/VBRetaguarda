VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPRINT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IMPRESSÃO"
   ClientHeight    =   2370
   ClientLeft      =   3045
   ClientTop       =   3330
   ClientWidth     =   4440
   Icon            =   "PRINT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   -120
      TabIndex        =   4
      Top             =   720
      Width           =   4695
      Begin VB.OptionButton optS 
         Caption         =   "&Sintético"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optA 
         Caption         =   "&Analítico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin MSMask.MaskEdBox maskINI 
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox maskFIM 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   1270
      ButtonWidth     =   2672
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Impressão"
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8760
         Top             =   120
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
               Picture         =   "PRINT.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRINT.frx":6DAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRINT.frx":7E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRINT.frx":8DF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PRINT.frx":9EFB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   DATA_INI = 0
   DATA_FIM = 0
   maskINI.PromptInclude = False
   maskFIM.PromptInclude = False
   maskINI.Text = ""
   maskFIM.Text = ""
   maskINI.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         optA.Value = False
         optS.Value = False
         maskINI.PromptInclude = False
         maskFIM.PromptInclude = False
         maskFIM.Text = ""
         maskINI.Text = ""
      Case "print"
         maskINI.PromptInclude = False
         maskFIM.PromptInclude = False
         If maskFIM.Text = "" Or maskINI.Text = "" Then
            MsgBox "Informe período corretamente."
            maskINI.SetFocus
            Exit Sub
         End If
         maskINI.PromptInclude = True
         maskFIM.PromptInclude = True
         DATA_FIM = maskFIM.Text
         DATA_INI = maskINI.Text
         Unload Me
   End Select
End Sub

Private Sub OPTA_Click()
   CONSULTA_A = 1
   cmdI.SetFocus
End Sub

Private Sub OPTS_Click()
   CONSULTA_A = 2
   cmdI.SetFocus
End Sub

Private Sub MaskFim_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdI.SetFocus
End Sub

Private Sub maskfim_LostFocus()
   maskFIM.PromptInclude = False
   If Not maskFIM.Text = "" Then
      maskFIM.PromptInclude = True
      If Not IsDate(maskFIM.Text) Then
         MsgBox "Data Informada Inválida !!!"
         maskFIM.SetFocus
         Exit Sub
      End If
   End If
   maskINI.PromptInclude = True
   maskFIM.PromptInclude = True
   If IsDate(maskINI.Text) And IsDate(maskFIM.Text) Then
      If CDate(maskINI.Text) > CDate(maskFIM.Text) Then
         MsgBox "Período Informado Inválido !!!"
         maskINI.SetFocus
         Exit Sub
      End If
   End If
   optA.SetFocus
End Sub

Private Sub maskini_GotFocus()
   maskINI.Mask = "##/##/####"
End Sub

Private Sub maskini_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      maskFIM.SetFocus
   End If
End Sub

Private Sub MaskFim_GotFocus()
   maskFIM.Mask = "##/##/####"
End Sub

Private Sub maskini_LostFocus()
   maskINI.PromptInclude = False
   If Not maskINI.Text = "" Then
      maskINI.PromptInclude = True
      If Not IsDate(maskINI.Text) Then
         MsgBox "Data Informada Inválida !!!"
         maskINI.SetFocus
      End If
   End If
End Sub
