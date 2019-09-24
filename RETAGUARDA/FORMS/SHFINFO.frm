VERSION 5.00
Object = "{EDF439C0-99E5-11CF-AFF3-004005100200}#8.0#0"; "PVMarq.ocx"
Begin VB.Form frmSHFINFO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Venda Lojinha"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   Icon            =   "SHFINFO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PVMarqueeLib.PVMarquee PVMarquee1 
      Height          =   615
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _Version        =   524288
      _ExtentX        =   10610
      _ExtentY        =   1085
      _StockProps     =   29
      Text            =   "SHF INFORMÁTICA E CONSULTORIA EM TI (62 8127-6360 / 8465-0219 / 9616-9569) "
      ForeColor       =   16711680
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Frame           =   5
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Text            =   "SHF INFORMÁTICA E CONSULTORIA EM TI (62 8127-6360 / 8465-0219 / 9616-9569) "
   End
   Begin VB.Label lblEJS 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   5280
   End
End
Attribute VB_Name = "frmSHFINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   CRITERIO_A = VB.App.Revision

   lblEJS.Caption = "Versão : " & VB.App.Major & "." & VB.App.Minor & "." & vbOKOnly & "." & vbInformation
End Sub

