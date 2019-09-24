VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCaixa 
   Caption         =   "Caixa Balcão"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CAIXA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotais 
      Height          =   855
      Left            =   60
      TabIndex        =   25
      Top             =   6000
      Width           =   7575
   End
   Begin MSComctlLib.Toolbar TBR 
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1270
      ButtonWidth     =   2487
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar Tela"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "imprimir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6480
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":5C12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":703A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":80C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":9263
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":A328
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":B002
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":2B7F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":2C97B
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CAIXA.frx":2CC95
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   60
      TabIndex        =   8
      Top             =   720
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Abertura/Lançamentos"
      TabPicture(0)   =   "CAIXA.frx":2ED17
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MSFlexGrid1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbCaixaAbre"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDtAbre"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbCaixaAbreAUX"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraBody"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtValorDig"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Fechamento"
      TabPicture(1)   =   "CAIXA.frx":2ED33
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5(3)"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "grdFaturados"
      Tab(1).Control(3)=   "grdPendentes"
      Tab(1).Control(4)=   "txtDtFecha"
      Tab(1).Control(5)=   "cmbCaixaFecha"
      Tab(1).Control(6)=   "cmbCaixaFechaAUX"
      Tab(1).Control(7)=   "cmdPendentes"
      Tab(1).Control(8)=   "cmdFaturados"
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdFaturados 
         Caption         =   "&Faturados"
         Height          =   375
         Left            =   -71160
         TabIndex        =   29
         Top             =   960
         Width           =   3615
      End
      Begin VB.CommandButton cmdPendentes 
         Caption         =   "&Pendentes"
         Height          =   375
         Left            =   -74880
         TabIndex        =   28
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtValorDig 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5760
         TabIndex        =   26
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbCaixaFechaAUX 
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
         Left            =   -73320
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbCaixaFecha 
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
         Left            =   -73320
         TabIndex        =   21
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtDtFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69720
         TabIndex        =   20
         Top             =   480
         Width           =   2175
      End
      Begin VB.Frame fraBody 
         Enabled         =   0   'False
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   7335
         Begin VB.ComboBox cmbModAUX 
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
            TabIndex        =   19
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbMod 
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
            Left            =   3240
            MousePointer    =   99  'Custom
            OLEDragMode     =   1  'Automatic
            TabIndex        =   18
            Top             =   720
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5760
            MaxLength       =   12
            TabIndex        =   5
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtSeq 
            Height          =   375
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   2
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtTIPO 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4080
            MaxLength       =   1
            TabIndex        =   4
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtMod 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   3
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtHistorico 
            Height          =   735
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   6
            Top             =   840
            Width           =   6015
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mod.:"
            Height          =   240
            Index           =   1
            Left            =   2040
            TabIndex        =   17
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor = "
            Height          =   240
            Left            =   4920
            TabIndex        =   16
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Seq.:"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo:(C/D)"
            Height          =   240
            Left            =   3030
            TabIndex        =   13
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.ComboBox cmbCaixaAbreAUX 
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
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDtAbre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cmbCaixaAbre 
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
         Left            =   1680
         TabIndex        =   0
         Top             =   600
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2175
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3836
         _Version        =   393216
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdPendentes 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   30
         Top             =   1440
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6376
         _Version        =   393216
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdFaturados 
         Height          =   3615
         Left            =   -71160
         TabIndex        =   31
         Top             =   1440
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6376
         _Version        =   393216
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70440
         TabIndex        =   23
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsável:"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Responsável:"
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   7320
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   7320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   7710
      DesignHeight    =   6930
   End
End
Attribute VB_Name = "frmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim TabCAIXA            As New ADODB.Recordset
   Dim CAIXADIA_ID_N       As Long
   Dim CAIXADIAITEM_ID_N   As Long
   Private LastRow         As Long ' Ultima linha em que se editou
   Private LastCol         As Long ' ultima coluna em que se editou
   Private ControlVisible  As Boolean
   Dim SQL_CORPO           As String
   Dim VLR_TOT_FAT_N       As Double

Private Sub Form_Load()

   MOSTRA_RESP
   CARREGA_FORMAPAGTO

   txtDtAbre.Text = Now
   cmbMod.Top = txtMod.Top    '1320
   'SSTab1.Tabs.

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      cmbCaixaAbre.Enabled = True
      cmbCaixaAbre.Text = ""
   End If
   TBR.Buttons(5).Visible = False
   TBR.Buttons(7).Visible = False

End Sub

Private Sub TBR_ButtonClick(ByVal Button As MSComctlLib.Button)

   Select Case Button.key
      Case "gravar"
         If SSTab1.Tab = 0 Then
            ABRE_CAIXA
            Else
               If SSTab1.Tab = 1 Then
                  FECHA_CAIXA
               End If
         End If
      Case "print"
      Case "limpar"
         CARREGA_FORMAPAGTO
         cmbCaixaAbre.Text = ""
         cmbCaixaAbreAUX.Text = ""
         cmbCaixaFecha.Text = ""
         cmbCaixaFechaAUX.Text = ""
         grdFaturados.Clear
         grdPendentes.Clear
         MSFlexGrid1.Clear
      Case "voltar"
         Unload Me
   End Select

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

   txtTotais.Text = ""
   grdPendentes.Clear
   cmdFaturados.Caption = "Faturados"
   If SSTab1.Tab = 0 Then
      TBR.Buttons(5).Caption = "Abrir Caixa"
      TBR.Buttons(5).Image = 9
      TBR.Buttons(5).Visible = False
      TBR.Buttons(7).Visible = False
      cmbCaixaAbre.SetFocus
   End If
   If SSTab1.Tab = 1 Then
      TBR.Buttons(5).Caption = "Fechar Caixa"
      TBR.Buttons(5).Image = 8
      TBR.Buttons(5).Visible = False
      TBR.Buttons(7).Visible = False
      txtDtFecha.Text = Now
      cmbCaixaFecha.SetFocus
   End If
   If SSTab1.Tab = 2 Then
      TBR.Buttons(5).Visible = False
      TBR.Buttons(7).Visible = True
   End If

End Sub

Private Sub cmbCaixaAbre_GotFocus()
'On Error GoTo ERRO_TRATA

   cmbCaixaAbre.SelStart = 0
   cmbCaixaAbre.SelLength = Len(cmbCaixaAbre.Text)
   cmbCaixaAbre.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCaixaAbre_GotFocus"
End Sub

Private Sub cmbCaixaAbre_Click()
'On Error GoTo ERRO_TRATA

   cmbCaixaAbreAUX.ListIndex = cmbCaixaAbre.ListIndex
   txtTotais.Text = ""
   If Trim(cmbCaixaAbre.Text) <> "" Then
      fraBody.Enabled = True
      Else: fraBody.Enabled = False
   End If

   txtSeq.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCaixaAbre_Click"
End Sub

Private Sub cmbCaixaAbre_LostFocus()
   cmbCaixaAbre.BackColor = &HFFFFFF
   If Trim(cmbCaixaAbreAUX.Text) <> "" Then _
      MOSTRA_MOVIMENTO_DIA cmbCaixaAbreAUX.Text, txtDtAbre.Text
End Sub

Private Sub txtseq_GotFocus()
'On Error GoTo ERRO_TRATA

   txtSeq.SelStart = 0
   txtSeq.SelLength = Len(txtSeq.Text)
   txtSeq.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSEQ_GotFocus"
End Sub

Private Sub txtseq_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      CAIXADIAITEM_ID_N = 0
      If Trim(txtSeq.Text) <> "" Then _
         If IsNumeric(txtSeq.Text) Then _
            CAIXADIAITEM_ID_N = 0 & txtSeq.Text

      If CAIXADIAITEM_ID_N = 0 Then
         CAIXADIAITEM_ID_N = 1

         If TabCAIXA.State = 1 Then _
            TabCAIXA.Close

         SQL = "select max(CAIXADIAITEM_ID) FROM CAIXADIA "
         SQL = SQL & " INNER JOIN CAIXADIAITEM "
         SQL = SQL & " ON CAIXADIA.CAIXADIA_ID = CAIXADIAITEM.CAIXADIA_ID "

         SQL = SQL & " where USUARIO_ID = " & cmbCaixaAbreAUX.Text
         SQL = SQL & " and dt_abertura >= '" & DMA(txtDtAbre.Text, "i") & "'"
         SQL = SQL & " and dt_abertura <= '" & DMA(txtDtAbre.Text, "f") & "'"

         TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCAIXA.EOF Then _
            If Not IsNull(TabCAIXA.Fields(0).Value) Then _
               CAIXADIAITEM_ID_N = 1 + TabCAIXA.Fields(0).Value
         If TabCAIXA.State = 1 Then _
            TabCAIXA.Close
      End If

      txtSeq.Text = CAIXADIAITEM_ID_N

      If CAIXADIAITEM_ID_N > 0 Then _
         MOSTRA_BODY

      txtMod.SetFocus

      KeyAscii = 0
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTSEQ_KeyPress"
End Sub

Private Sub txtseq_LostFocus()
   txtSeq.BackColor = &HFFFFFF
   MOSTRA_SEQUENCIA
End Sub

Private Sub txtMOD_GotFocus()
'On Error GoTo ERRO_TRATA

   txtMod.SelStart = 0
   txtMod.SelLength = Len(txtMod.Text)
   txtMod.BackColor = &HC0FFFF

   cmbMod.Text = "--- Selecione ---"
   cmbMod.Visible = True

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtMOD_GotFocus"
End Sub

Private Sub txtmod_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      If Trim(txtMod.Text) = "" Then _
         txtMod.Text = 1

      If TabDESCR.State = 1 Then _
         TabDESCR.Close
      SQL = "select * from FORMAPAGTO WITH (NOLOCK) "
      SQL = SQL & " where formapagto_id = " & txtMod.Text
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and status = 'true' "
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then
         txtMod.Text = TabDESCR!FORMAPAGTO_ID
         txtHistorico.Text = TabDESCR!DESCRICAO
      End If
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
      
      KeyAscii = 0

      cmbMod.Visible = False
      txtTIPO.SetFocus
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtmod_KeyPress"
End Sub

Private Sub txtMOD_LostFocus()
   txtMod.BackColor = &HFFFFFF
End Sub

Private Sub cmbMOD_Click()
'On Error GoTo ERRO_TRATA

   cmbModAUX.ListIndex = cmbMod.ListIndex

   If Trim(cmbModAUX.Text) <> "" Then
      txtMod.Text = cmbModAUX.Text
      txtTIPO.SetFocus
      cmbMod.Visible = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbMOD_Click"
End Sub

Private Sub txtTIPO_GotFocus()
'On Error GoTo ERRO_TRATA

   txtTIPO.SelStart = 0
   txtTIPO.SelLength = Len(txtTIPO.Text)
   txtTIPO.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTIPO_GotFocus"
End Sub

Private Sub txtTipo_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtTIPO.Text = "D" Or txtTIPO.Text = "C" Then _
         txtValor.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtTipo_KeyPress"
End Sub

Private Sub txtTIPO_LostFocus()
   txtTIPO.BackColor = &HFFFFFF
End Sub

Private Sub txtTotais_GotFocus()
   txtSeq.SetFocus
End Sub

Private Sub txtvalor_GotFocus()
   txtValor.SelStart = 0
   txtValor.SelLength = Len(txtValor.Text)
   txtValor.BackColor = &HC0FFFF
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      txtHistorico.SetFocus
      Else
         If KeyAscii = 8 Or KeyAscii = 44 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TXTVALOR_KeyPress"
End Sub

Private Sub txtValor_LostFocus()
   txtValor.BackColor = &HFFFFFF
   If Trim(txtValor.Text) <> "" Then
      VALOR_ITEM_N = 0 & txtValor.Text
      txtValor.Text = "" & Format(VALOR_ITEM_N, strFormatacao2Digitos)
   End If
End Sub

Private Sub txthistorico_GotFocus()
'On Error GoTo ERRO_TRATA

   txtHistorico.SelStart = 0
   txtHistorico.SelLength = Len(txtHistorico.Text)
   txtHistorico.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txthistorico_gotFocus"
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      GRAVA_REGISTRO
      MOSTRA_MOVIMENTO_DIA cmbCaixaAbreAUX.Text, txtDtAbre.Text

      txtSeq.SetFocus
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtHISTORICO_KeyPress"
End Sub

Private Sub txthistorico_LostFocus()
   txtHistorico.BackColor = &HFFFFFF
End Sub

Private Sub cmbCaixaFECHA_GotFocus()
'On Error GoTo ERRO_TRATA

   cmbCaixaFecha.SelStart = 0
   cmbCaixaFecha.SelLength = Len(cmbCaixaFecha.Text)
   cmbCaixaFecha.BackColor = &HC0FFFF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCaixaFECHA_GotFocus"
End Sub

Private Sub cmbCaixaFECHA_Click()
'On Error GoTo ERRO_TRATA

   cmbCaixaFechaAUX.ListIndex = cmbCaixaFecha.ListIndex
   txtTotais.Text = ""
   If Trim(cmbCaixaFecha.Text) <> "" Then
      If Trim(cmbCaixaFechaAUX.Text) <> "" Then _
         MOSTRA_MOVIMENTO_DIA cmbCaixaFechaAUX.Text, txtDtFecha.Text
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmbCaixaFECHA_Click"
End Sub

Private Sub cmbCaixaFECHA_LostFocus()
   cmbCaixaFecha.BackColor = &HFFFFFF
End Sub

Private Sub txtValorDig_GotFocus()
'On Error GoTo ERRO_TRATA

   txtValorDig.SelStart = 0
   txtValorDig.SelLength = Len(txtValorDig)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtValorDig_GotFocus"
End Sub

Private Sub cmdPendentes_Click()
'On Error GoTo ERRO_TRATA

   SQL = ""
   SQL_CORPO = ""
   If Trim(cmbCaixaFechaAUX.Text) <> "" Then
      SQL = "select pedido_id as Pedido, NOME_CLIENTE as Cliente, dt_req as Data, VALOR_TOTAL as Valor, pedido.status as ST_PEDIDO "
      SQL = SQL & " from PEDIDO "
      SQL = SQL & " INNER JOIN VENDEDOR "
      SQL = SQL & " ON PEDIDO.VENDEDOR_ID = VENDEDOR.VENDEDOR_ID "
      SQL = SQL & " INNER JOIN USUARIO "
      SQL = SQL & " ON VENDEDOR.PESSOA_ID = USUARIO.PESSOA_ID"

      SQL = SQL & " where pedido.STATUS < 3 "
      SQL = SQL & " and Usuario.USUARIO_ID = " & Trim(cmbCaixaFechaAUX.Text)

      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL = SQL & " and dt_req >= '" & DMA(txtDtFecha.Text, "i") & "'"
      SQL = SQL & " and dt_req <= '" & DMA(txtDtFecha.Text, "f") & "'"

      SETA_GRID_FECHA_PEND SQL, cmbCaixaFechaAUX.Text
      MOSTRA_TOTAIS cmbCaixaFechaAUX.Text
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdPendentes_Click"
End Sub

Private Sub cmdFaturados_Click()
'On Error GoTo ERRO_TRATA

   SQL = ""
   SQL_CORPO = ""

   If Trim(cmbCaixaFechaAUX.Text) <> "" Then
      SQL = "select pedido_id as Pedido, NOME_CLIENTE as Cliente, dt_req as Data, "
      SQL = SQL & " VALOR_TOTAL as Valor, pedido.status as ST_PEDIDO "

      SQL_CORPO = SQL_CORPO & " from PEDIDO "
      SQL_CORPO = SQL_CORPO & " INNER JOIN VENDEDOR "
      SQL_CORPO = SQL_CORPO & " ON PEDIDO.VENDEDOR_ID = VENDEDOR.VENDEDOR_ID "
      SQL_CORPO = SQL_CORPO & " INNER JOIN USUARIO "
      SQL_CORPO = SQL_CORPO & " ON VENDEDOR.PESSOA_ID = USUARIO.PESSOA_ID"

      SQL_CORPO = SQL_CORPO & " where pedido.status in (3,5,7) "
      SQL_CORPO = SQL_CORPO & " and Usuario.USUARIO_ID = " & Trim(cmbCaixaFechaAUX.Text)

      SQL_CORPO = SQL_CORPO & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SQL_CORPO = SQL_CORPO & " and dt_req >= '" & DMA(txtDtFecha.Text, "i") & "'"
      SQL_CORPO = SQL_CORPO & " and dt_req <= '" & DMA(txtDtFecha.Text, "f") & "'"

      'SETA_GRID_FECHA_FAT SQL & " " & SQL_CORPO, cmbCaixaFechaAUX.Text

      SqL2 = "SELECT DESCRICAO, sum(QTD_PEDIDA * VALOR_ITEM) AS Total"

      SqL2 = SqL2 & " FROM PEDIDO "
      SqL2 = SqL2 & " INNER JOIN PEDIDOITEM "
      SqL2 = SqL2 & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
      SqL2 = SqL2 & " INNER JOIN PEDIDOFATURA "
      SqL2 = SqL2 & " ON PEDIDO.PEDIDO_ID = PEDIDOFATURA.PEDIDO_ID "
      SqL2 = SqL2 & " INNER JOIN TIPOVENDA "
      SqL2 = SqL2 & " ON PEDIDOFATURA.TIPOVENDA_ID = TIPOVENDA.TIPOVENDA_ID"


      SqL2 = SqL2 & " where pedido.status in (3,5,7) "
      SqL2 = SqL2 & " and USUARIO_ID = " & Trim(cmbCaixaFechaAUX.Text)

      SqL2 = SqL2 & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      SqL2 = SqL2 & " and dt_req >= '" & DMA(txtDtFecha.Text, "i") & "'"
      SqL2 = SqL2 & " and dt_req <= '" & DMA(txtDtFecha.Text, "f") & "'"

      SqL2 = SqL2 & " group by DESCRICAO"

SETA_GRID_FECHA_FAT SqL2, cmbCaixaFechaAUX.Text

      MOSTRA_TOTAIS cmbCaixaFechaAUX.Text

      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      SQL = "select sum(VALOR_TOTAL) "
      TabCAIXA.Open SQL & " " & SQL_CORPO, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCAIXA.EOF Then
         If Trim(txtTotais.Text) <> "" Then
            txtTotais.Text = Trim(txtTotais.Text) & " ; Pedidos = " & Format(TabCAIXA.Fields(0).Value, strFormatacao2Digitos)
            Else: txtTotais.Text = " Pedidos = " & Format(TabCAIXA.Fields(0).Value, strFormatacao2Digitos)
         End If
      End If
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      txtTotais.Refresh
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdFaturados_Click"
End Sub

Private Sub txtValorDig_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ERRO_TRATA

   Select Case KeyCode
      Case vbKeyEscape
         MSFlexGrid1.SetFocus
      Case vbKeyUp
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
               txtSeq.SetFocus
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

         'SQL = SQL & " where PEDIDO.PEDIDO_ID = " & txtPedido.Text
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
      txtSeq.SetFocus
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
Private Sub MSFlexGrid1_DblClick()
'On Error GoTo ERRO_TRATA

   'editar ao clicar duas vezes
   LastRow = MSFlexGrid1.Row
   LastCol = MSFlexGrid1.Col

   ExibirCelula

   'txtseq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 0)
   'txtSeq.Text = "" & MSFlexGrid1.TextMatrix(LastRow, 11)

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
            If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9)) Then
               CAIXADIAITEM_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9)
               CAIXADIA_ID_N = 0 & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8)
               Acao_N = 3
               SQL = "spCaixaItem " & Acao_N & "," & CAIXADIA_ID_N & "," & CAIXADIAITEM_ID_N & "," & 1 & ",'" & tpMOEDA(1) & "','" & Trim(1) & "','" & Trim("1") & "'"
               CONECTA_RETAGUARDA.Execute "EXEC " & SQL
               MOSTRA_MOVIMENTO_DIA cmbCaixaAbreAUX.Text, txtDtAbre.Text
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
'================================
'================================
Private Sub MOSTRA_RESP()
'On Error GoTo ERRO_TRATA

   cmbCaixaAbre.Enabled = True
   cmbCaixaAbreAUX.Enabled = True
   cmbCaixaAbre.Clear
   cmbCaixaAbreAUX.Clear
   cmbCaixaFecha.Clear
   cmbCaixaFechaAUX.Clear

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "SELECT USUARIO.USUARIO_ID, USUARIO.EMPRESA_ID, USUARIO.NOME, USUARIO.CPF, "
   SQL = SQL & " USUARIO.TIPO, USUARIO.STATUS, USUARIO.LOGON, USUARIO.PESSOA_ID, "
   SQL = SQL & " Usuario.FUNCIONARIO , Vendedor.EQUIPE_ID, Vendedor.vendedor_ID"
   SQL = SQL & " FROM USUARIO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN VENDEDOR WITH (NOLOCK)"
   SQL = SQL & " ON USUARIO.PESSOA_ID = VENDEDOR.PESSOA_ID"

   SQL = SQL & " where Usuario.TIPO = 7 "
   SQL = SQL & " And Usuario.Status = 1"
   SQL = SQL & " and VENDEDOR.status = 'A' "

   If TIPO_USUARIO = 4 Or TIPO_USUARIO = 5 Then
      Else: SQL = SQL & " And Usuario.usuario_id = " & USUARIO_ID_N
   End If

   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabUSU.EOF
      cmbCaixaAbre.AddItem Trim(TabUSU.Fields("nome").Value) & "-" & TabUSU.Fields("usuario_id").Value
      cmbCaixaAbreAUX.AddItem TabUSU.Fields("usuario_id").Value

      cmbCaixaFecha.AddItem Trim(TabUSU.Fields("nome").Value) & "-" & TabUSU.Fields("usuario_id").Value
      cmbCaixaFechaAUX.AddItem TabUSU.Fields("usuario_id").Value

      If TabUSU.Fields("usuario_id").Value = USUARIO_ID_N Then
         cmbCaixaAbre.Text = "" & Trim(TabUSU.Fields("nome").Value)
         cmbCaixaAbreAUX.Text = "" & TabUSU.Fields("usuario_id").Value
      End If
      If TabUSU.Fields("usuario_id").Value = USUARIO_ID_N Then
         cmbCaixaFecha.Text = "" & Trim(TabUSU.Fields("nome").Value)
         cmbCaixaFechaAUX.Text = "" & TabUSU.Fields("usuario_id").Value
      End If

      TabUSU.MoveNext
   Wend
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If Trim(cmbCaixaAbre.Text) <> "" Then
      fraBody.Enabled = True
      Else: fraBody.Enabled = False
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_RESP"
End Sub

Sub LIMPA_BODY()
'On Error GoTo ERRO_TRATA

   txtSeq.Text = ""
   txtMod.Text = ""
   txtTIPO.Text = ""
   txtValor.Text = ""
   txtHistorico.Text = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_BODY"
End Sub

Sub MOSTRA_BODY()
'On Error GoTo ERRO_TRATA

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select CAIXADIA_ID from CAIXADIA WITH (NOLOCK) "
   SQL = SQL & " where usuario_id = " & Trim(cmbCaixaAbreAUX.Text)
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and dt_abertura >= '" & DMA(txtDtAbre.Text, "i") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(txtDtAbre.Text, "f") & "'"
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then _
      CAIXADIA_ID_N = 0 & Trim(TabCAIXA.Fields(0).Value)

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select * from CAIXADIAITEM WITH (NOLOCK) "
   SQL = SQL & " where CAIXADIA_ID = " & CAIXADIA_ID_N
   SQL = SQL & " and CAIXADIAITEM_ID = " & txtSeq.Text
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then
      txtMod.Text = "" & TabCAIXA.Fields("FORMAPAGTO_ID").Value
      txtTIPO.Text = "" & TabCAIXA.Fields("TIPO").Value
      txtValor.Text = "" & Format(TabCAIXA.Fields("VALOR").Value, strFormatacao2Digitos)
      txtHistorico.Text = "" & TabCAIXA.Fields("HISTORICO").Value
   End If
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_BODY"
End Sub

Sub CARREGA_FORMAPAGTO()
'On Error GoTo ERRO_TRATA

   cmbMod.Clear
   cmbModAUX.Clear

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select formapagto_id,descricao from FORMAPAGTO WITH (NOLOCK) "
   SQL = SQL & " where status = 1 "
   SQL = SQL & " and contab_balcao = 1 "
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      cmbMod.AddItem Trim(TabConsulta.Fields("descricao").Value) & "-" & TabConsulta.Fields("formapagto_id").Value
      cmbModAUX.AddItem TabConsulta.Fields("formapagto_id").Value

      TabConsulta.MoveNext
   Wend
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CARREGA_FORMAPAGTO"
End Sub

Sub MOSTRA_MOVIMENTO_DIA(Codg_Usu_N As Long, DtMov_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(Codg_Usu_N) <> "" Then
      SQL = "SELECT Historico, FORMAPAGTO.DESCRICAO as PAGTO, CAIXADIAITEM.TIPO AS TPPAGTO, CAIXADIAITEM.Valor, "
      SQL = SQL & " DT_ABERTURA as DtAbertura,DT_FECHAMENTO as DtFechamento "

      SQL = SQL & " ,NUMERO_CAIXA_CPU, CAIXADIAITEM.FORMAPAGTO_ID, CAIXADIA.CAIXADIA_ID , CAIXADIAITEM_ID"

      SQL = SQL & " FROM CAIXADIA WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN CAIXADIAITEM WITH (NOLOCK)"
      SQL = SQL & " ON CAIXADIA.CAIXADIA_ID = CAIXADIAITEM.CAIXADIA_ID "
      SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " ON CAIXADIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
      SQL = SQL & " Inner Join USUARIO WITH (NOLOCK)"
      SQL = SQL & " ON CAIXADIA.USUARIO_ID = USUARIO.USUARIO_ID"

      SQL = SQL & " where CAIXADIA.USUARIO_ID = " & Codg_Usu_N
      SQL = SQL & " and dt_abertura >= '" & DMA(DtMov_A, "i") & "'"
      SQL = SQL & " and dt_abertura <= '" & DMA(DtMov_A, "f") & "'"

      SQL = SQL & " order by CAIXADIAITEM.TIPO desc"

      SETA_GRID SQL, Codg_Usu_N

      If CHECAR_CAIXA(Trim(Codg_Usu_N), DMA(txtDtAbre.Text)) = False Then
         TBR.Buttons(5).Caption = "Abrir Caixa"
         TBR.Buttons(5).Image = 9
         TBR.Buttons(5).Visible = True
         fraBody.Enabled = False

         If SSTab1.Tab = 1 Then
            If Trim(cmbCaixaFecha.Text) <> "" Then
               MSFlexGrid1.Clear

               cmbCaixaAbre.Text = cmbCaixaFecha.Text
               cmbCaixaAbreAUX.Text = cmbCaixaFechaAUX.Text
               SSTab1.Tab = 0
            End If
         End If

         cmbCaixaAbre.SetFocus
         Else
            If SSTab1.Tab = 1 Then
               MSFlexGrid1.Clear
               TBR.Buttons(5).Caption = "Fechar Caixa"
               TBR.Buttons(5).Image = 8
               TBR.Buttons(5).Visible = True
               TBR.Buttons(7).Visible = False
               txtDtFecha.Text = Now
               Else
                  TBR.Buttons(5).Visible = False
                  fraBody.Enabled = True
                  txtSeq.SetFocus
            End If
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_MOVIMENTO_DIA"
End Sub

Sub SETA_GRID(strSQL As String, Codg_Usu_N As Long)
'On Error GoTo ERRO_TRATA

   If Codg_Usu_N <= 0 Then _
      Exit Sub

   Dim TabGrid As New ADODB.Recordset
   Dim Coluna, Linha, Largura_Campo

   CONT_N = 0
   VALOR_ITEM_N = 0
   VALOR_TOTAL_N = 0

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
'MsgBox strSQL
   TabGrid.Open strSQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGrid.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      MSFlexGrid1.Rows = 2
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
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

         For Coluna = 0 To TabGrid.Fields.Count - 1
            'If Coluna = 3 Or Coluna = 7 Then
            If Coluna = 3 Then
               MSFlexGrid1.TextMatrix(Linha, Coluna) = Format(TabGrid.Fields(Coluna).Value, strFormatacao3Digitos)
               Else: MSFlexGrid1.TextMatrix(Linha, Coluna) = "" & Trim(TabGrid.Fields(Coluna).Value)
            End If

            If TabGrid.Fields("TPPAGTO").Value = "D" Then
               MSFlexGrid1.Row = Linha
               MSFlexGrid1.Col = Coluna
               'flex_tst.Text = "Bold Font"
               'flex_tst.CellFontBold = True
               'flex_tst.CellForeColor = vbRed
               MSFlexGrid1.CellForeColor = vbRed ' &H4000&   '&H40&
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabGrid.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabGrid.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo
         Next Coluna

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
   

'Histórico
      MSFlexGrid1.ColWidth(0) = 6000
      MSFlexGrid1.ColAlignment(0) = 0

'PAGTO
      MSFlexGrid1.ColWidth(1) = 3000
      MSFlexGrid1.ColAlignment(1) = 0

'TPPAGTO
      MSFlexGrid1.ColWidth(2) = 1000
      MSFlexGrid1.ColAlignment(2) = 0

'Valor
      MSFlexGrid1.ColWidth(3) = 3000
      MSFlexGrid1.ColAlignment(3) = 7

'DT_ABERTURA
      MSFlexGrid1.ColWidth(4) = 3100
      MSFlexGrid1.ColAlignment(4) = 7

'DT_FECHAMENTO
      MSFlexGrid1.ColWidth(5) = 3100
      MSFlexGrid1.ColAlignment(5) = 7

'NUMERO_CAIXA_CPU
      MSFlexGrid1.ColWidth(6) = 1

'FORMAPAGTO_ID
      MSFlexGrid1.ColWidth(7) = 1

'CAIXADIA.CAIXADIA_ID
      MSFlexGrid1.ColWidth(8) = 1000
      MSFlexGrid1.ColWidth(8) = 1

'CAIXADIAITEM_ID
      MSFlexGrid1.ColWidth(9) = 1000
      MSFlexGrid1.ColWidth(9) = 1

   If TabGrid.State = 1 Then _
      TabGrid.Close

   MOSTRA_TOTAIS Codg_Usu_N

   DoEvents

   MSFlexGrid1.Visible = True
   
   End If
'CellFontName        - Define o nome da fonte para uma célula
'CellFontSize        - Define o tamanho da fonte para a célula
'CellFontBold        - Define se a fonte aparece em negrito.
'CellFontItalic      - Define se a fonte aparece em itálico.
'CellFontUnderline   - Define se a fonte aparece sublinhada.

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID"
End Sub

Sub MOSTRA_TOTAIS(Codg_Usu_N As Long)
'On Error GoTo ERRO_TRATA

   'lstTotais.ListItems.Clear
   CONT_N = 0
   txtTotais.Text = ""

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select descricao,tipo, sum(valor) as VLR "
   SQL = SQL & " FROM CAIXADIA "
   SQL = SQL & " INNER JOIN CAIXADIAITEM "
   SQL = SQL & " ON CAIXADIA.CAIXADIA_ID = CAIXADIAITEM.CAIXADIA_ID "
   SQL = SQL & " INNER JOIN FORMAPAGTO "
   SQL = SQL & " ON CAIXADIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"

   SQL = SQL & " where USUARIO_ID = " & Codg_Usu_N
   SQL = SQL & " and dt_abertura >= '" & DMA(txtDtAbre.Text, "i") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(txtDtAbre.Text, "f") & "'"

   SQL = SQL & " group by descricao,tipo"

   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCAIXA.EOF
      If Trim(txtTotais.Text) = "" Then
         If Trim(TabCAIXA.Fields("tipo").Value) = "C" Then
            txtTotais.Text = txtTotais.Text & Trim(TabCAIXA.Fields("descricao").Value) & " = " & Format(TabCAIXA.Fields("vlr").Value, strFormatacao2Digitos)
            Else: txtTotais.Text = txtTotais.Text & Trim(TabCAIXA.Fields("descricao").Value) & " = -" & Format(TabCAIXA.Fields("vlr").Value, strFormatacao2Digitos)
         End If
         Else
            If Trim(TabCAIXA.Fields("tipo").Value) = "C" Then
               txtTotais.Text = txtTotais.Text & " ; " & Trim(TabCAIXA.Fields("descricao").Value) & " = " & Format(TabCAIXA.Fields("vlr").Value, strFormatacao2Digitos)
               Else: txtTotais.Text = txtTotais.Text & " ; " & Trim(TabCAIXA.Fields("descricao").Value) & " = -" & Format(TabCAIXA.Fields("vlr").Value, strFormatacao2Digitos)
            End If
      End If

      CONT_N = CONT_N + 1

      TabCAIXA.MoveNext
   Wend
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ABRE_CAIXA"
End Sub

Sub MOSTRA_SEQUENCIA()
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "SELECT HISTORICO, caixadiaitem.formapagto_id as forma_id, "
   SQL = SQL & " CAIXADIAITEM.TIPO AS TPPAGTO, CAIXADIAITEM.valor as VlrDig "

   SQL = SQL & " FROM CAIXADIA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN CAIXADIAITEM WITH (NOLOCK)"
   SQL = SQL & " ON CAIXADIA.CAIXADIA_ID = CAIXADIAITEM.CAIXADIA_ID "
   SQL = SQL & " INNER JOIN FORMAPAGTO WITH (NOLOCK)"
   SQL = SQL & " ON CAIXADIAITEM.FORMAPAGTO_ID = FORMAPAGTO.FORMAPAGTO_ID"
   SQL = SQL & " Inner Join USUARIO WITH (NOLOCK)"
   SQL = SQL & " ON CAIXADIA.USUARIO_ID = USUARIO.USUARIO_ID"

   SQL = SQL & " where CAIXADIAITEM.CAIXADIA_ID = " & CAIXADIA_ID_N
   SQL = SQL & " and CAIXADIAITEM.CAIXADIAITEM_ID = " & CAIXADIAITEM_ID_N

   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabConsulta.EOF Then
      txtMod.Text = "" & TabConsulta.Fields("forma_id").Value
      txtTIPO.Text = "" & TabConsulta.Fields("TPPAGTO").Value
      txtValor.Text = "" & Format(TabConsulta.Fields("VlrDig").Value, strFormatacao2Digitos)
      txtHistorico.Text = "" & TabConsulta.Fields("historico").Value
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "MOSTRA_SEQUENCIA"
End Sub

Sub GRAVA_REGISTRO()
'On Error GoTo ERRO_TRATA

   If Trim(cmbCaixaAbreAUX.Text) = "" Then
      MsgBox "Informar Responsável."
      cmbCaixaAbre.SetFocus
      Exit Sub
   End If
   CAIXADIAITEM_ID_N = 0 & txtSeq.Text
   If CAIXADIAITEM_ID_N <= 0 Then
      MsgBox "Informar Seqüência."
      txtSeq.SetFocus
      Exit Sub
   End If
   If Trim(txtMod.Text) = "" Then
      MsgBox "Informar Pagto."
      txtTIPO.SetFocus
      Exit Sub
   End If
   If Trim(txtTIPO.Text) = "" Then
      MsgBox "Informar Débito/Credito."
      txtTIPO.SetFocus
      Exit Sub
   End If
   VALOR_ITEM_N = 0 & txtValor
   If VALOR_ITEM_N <= 0 Then
      MsgBox "Informar Valor."
      txtValor.SetFocus
      Exit Sub
   End If
   If Trim(txtHistorico.Text) = "" Then
      MsgBox "Informar Histórico."
      txtHistorico.SetFocus
      Exit Sub
   End If

   CAIXADIA_ID_N = 0
   Acao_N = 0

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select CAIXADIA_ID from CAIXADIA WITH (NOLOCK) "
   SQL = SQL & " where usuario_id = " & Trim(cmbCaixaAbreAUX.Text)
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and dt_abertura >= '" & DMA(txtDtAbre.Text, "i") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(txtDtAbre.Text, "f") & "'"
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then _
      CAIXADIA_ID_N = 0 & Trim(TabCAIXA.Fields(0).Value)
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   If CAIXADIA_ID_N <= 0 Then
      Acao_N = 1
      txtDtFecha.Text = ""
      Else
         If SSTab1.Tab = 1 Then
            Acao_N = 2
            txtDtFecha.Text = Now
         End If
   End If
   If Acao_N > 0 And Acao_N < 4 Then
      SQL = "spCaixa " & Acao_N & "," & CAIXADIA_ID_N & "," & Trim(cmbCaixaAbreAUX.Text) & "," & ESTABELECIMENTO_ID_N & "," & NUMERO_CAIXA_CPU & ",'" & Trim(txtDtAbre.Text) & "','" & Trim(txtDtFecha.Text) & "'"
      CONECTA_RETAGUARDA.Execute "EXEC " & SQL
   End If

'=============================

   SQL = "select CAIXADIAITEM_ID from CAIXADIAITEM WITH (NOLOCK) "
   SQL = SQL & " where CAIXADIA_ID = " & CAIXADIA_ID_N
   SQL = SQL & " and CAIXADIAITEM_ID = " & CAIXADIAITEM_ID_N
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then
      CAIXADIAITEM_ID_N = 0 & Trim(TabCAIXA.Fields(0).Value)
      Acao_N = 2
      Else: Acao_N = 1
   End If
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

'pegando id caixa pra gravar no item
   SQL = "select CAIXADIA_ID from CAIXADIA WITH (NOLOCK) "
   SQL = SQL & " where usuario_id = " & Trim(cmbCaixaAbreAUX.Text)
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and dt_abertura >= '" & DMA(txtDtAbre.Text, "i") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(txtDtAbre.Text, "f") & "'"
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then _
      CAIXADIA_ID_N = 0 & Trim(TabCAIXA.Fields(0).Value)
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   If Acao_N > 0 And Acao_N < 4 Then
      SQL = "spCaixaItem " & Acao_N & "," & CAIXADIA_ID_N & "," & CAIXADIAITEM_ID_N & "," & txtMod.Text & ",'" & tpMOEDA(VALOR_ITEM_N) & "','" & Trim(txtTIPO.Text) & "','" & Trim(txtHistorico.Text) & "'"
      CONECTA_RETAGUARDA.Execute "EXEC " & SQL
   End If
   LIMPA_BODY

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GRAVA_REGISTRO"
End Sub

Sub ABRE_CAIXA()
'On Error GoTo ERRO_TRATA

   If Trim(cmbCaixaAbreAUX.Text) = "" Then
      MsgBox "Informar Responsável."
      cmbCaixaAbre.SetFocus
      Exit Sub
   End If

   CAIXADIA_ID_N = 0
   Acao_N = 0

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select CAIXADIA_ID from CAIXADIA WITH (NOLOCK) "
   SQL = SQL & " where usuario_id = " & Trim(cmbCaixaAbreAUX.Text)
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and dt_abertura >= '" & DMA(txtDtAbre.Text, "i") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(txtDtAbre.Text, "f") & "'"
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then
      CAIXADIA_ID_N = 0 & Trim(TabCAIXA.Fields(0).Value)
      Acao_N = 2
      Else
         CAIXADIA_ID_N = 0 & MAX_ID("CAIXADIA_ID", "CAIXADIA", "", "", "", "")
         Acao_N = 1
         txtDtFecha.Text = ""
   End If
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   If Acao_N = 1 Then
      SQL = "spCaixa " & Acao_N & "," & CAIXADIA_ID_N & "," & Trim(cmbCaixaAbreAUX.Text) & "," & ESTABELECIMENTO_ID_N & "," & NUMERO_CAIXA_CPU & ",'" & Trim(txtDtAbre.Text) & "','" & Trim(txtDtFecha.Text) & "'"
      CONECTA_RETAGUARDA.Execute "EXEC " & SQL
   End If
   cmbCaixaAbre.SetFocus
   If CAIXADIA_ID_N > 0 Then
      MsgBox "Caixa aberto com sucesso !!!"
      fraBody.Enabled = True
      txtSeq.SetFocus
      Else: MsgBox "Erro na abertura do caixa."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ABRE_CAIXA"
End Sub

Sub FECHA_CAIXA()
'On Error GoTo ERRO_TRATA

   If Trim(cmbCaixaFechaAUX.Text) = "" Then
      MsgBox "Informar Responsável."
      cmbCaixaFecha.SetFocus
      Exit Sub
   End If

   CAIXADIA_ID_N = 0
   Acao_N = 2

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select CAIXADIA_ID from CAIXADIA WITH (NOLOCK) "
   SQL = SQL & " where usuario_id = " & Trim(cmbCaixaFechaAUX.Text)
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and dt_abertura >= '" & DMA(txtDtFecha.Text, "i") & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(txtDtFecha.Text, "f") & "'"
   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCAIXA.EOF Then _
      CAIXADIA_ID_N = 0 & Trim(TabCAIXA.Fields(0).Value)
   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   If Acao_N = 2 Then
      SQL = "spCaixa " & Acao_N & "," & CAIXADIA_ID_N & "," & Trim(cmbCaixaFechaAUX.Text) & "," & ESTABELECIMENTO_ID_N & "," & 1 & ",'" & Trim(txtDtAbre.Text) & "','" & Trim(txtDtFecha.Text) & "'"
      CONECTA_RETAGUARDA.Execute "EXEC " & SQL
   End If

   If CAIXADIA_ID_N > 0 Then
      MsgBox "Caixa Fechado com sucesso !!!"
      SSTab1.Tab = 0
      MSFlexGrid1.Clear
      Else: MsgBox "Erro Fechamento do caixa."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FECHA_CAIXA"
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

Sub SETA_GRID_FECHA_PEND(strSQL As String, Codg_Usu_N As Long)
'On Error GoTo ERRO_TRATA

   If Codg_Usu_N <= 0 Then _
      Exit Sub

   Dim TabGrid As New ADODB.Recordset
   Dim Coluna, Linha, Largura_Campo

   CONT_N = 0
   VALOR_ITEM_N = 0
   VALOR_TOTAL_N = 0

   grdPendentes.Clear
   grdPendentes.Visible = False
   grdPendentes.Gridlines = flexGridFlat
   grdPendentes.FixedRows = 1
   grdPendentes.FixedCols = 1
   grdPendentes.ScrollBars = flexScrollBarBoth
   grdPendentes.AllowUserResizing = flexResizeColumns

   If TabGrid.State = 1 Then _
      TabGrid.Close

   TabGrid.Open strSQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGrid.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      grdPendentes.Rows = 2
      grdPendentes.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      grdPendentes.Rows = 1
      grdPendentes.Cols = TabGrid.Fields.Count

      ReDim largura_coluna(0 To TabGrid.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabGrid.Fields.Count - 1
         grdPendentes.TextMatrix(0, Coluna) = Trim(TabGrid.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabGrid.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabGrid.EOF
         grdPendentes.Rows = grdPendentes.Rows + 1

         For Coluna = 0 To TabGrid.Fields.Count - 1
            If Coluna = 3 Then
               grdPendentes.TextMatrix(Linha, Coluna) = Format(TabGrid.Fields(Coluna).Value, strFormatacao3Digitos)
               Else: grdPendentes.TextMatrix(Linha, Coluna) = "" & Trim(TabGrid.Fields(Coluna).Value)
            End If

            If TabGrid.Fields("ST_PEDIDO").Value < 3 Then
               grdPendentes.Row = Linha
               grdPendentes.Col = Coluna
               grdPendentes.CellForeColor = vbRed ' &H4000&   '&H40&
            End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabGrid.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabGrid.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo
         Next Coluna

         TabGrid.MoveNext
         Linha = Linha + 1
      Loop

      If TabGrid.State = 1 Then _
         TabGrid.Close

      'define a largura das colunas do grid
      For Coluna = 0 To grdPendentes.Cols - 1
         grdPendentes.ColWidth(Coluna) = largura_coluna(Coluna) + 240
      Next Coluna

      grdPendentes.ColWidth(0) = 0
      grdPendentes.Refresh

      grdPendentes.BackColor = vbWhite
      grdPendentes.ForeColor = vbBlue
   
'pedido
      grdPendentes.ColWidth(0) = 2000
      grdPendentes.ColAlignment(0) = 0

'cliente
      grdPendentes.ColWidth(1) = 3000
      grdPendentes.ColAlignment(1) = 0

'dt
      grdPendentes.ColWidth(2) = 2000
      grdPendentes.ColAlignment(2) = 0

'Valor
      grdPendentes.ColWidth(3) = 2000
      grdPendentes.ColAlignment(3) = 7

'situacao
      grdPendentes.ColWidth(4) = 100
      grdPendentes.ColAlignment(4) = 0

      MOSTRA_TOTAIS Codg_Usu_N

      DoEvents

      grdPendentes.Visible = True
   
   End If
'CellFontName        - Define o nome da fonte para uma célula
'CellFontSize        - Define o tamanho da fonte para a célula
'CellFontBold        - Define se a fonte aparece em negrito.
'CellFontItalic      - Define se a fonte aparece em itálico.
'CellFontUnderline   - Define se a fonte aparece sublinhada.

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_FECHA_PEND"
End Sub

Sub SETA_GRID_FECHA_FAT(strSQL As String, Codg_Usu_N As Long)
'On Error GoTo ERRO_TRATA

   If Codg_Usu_N <= 0 Then _
      Exit Sub

   Dim TabGrid As New ADODB.Recordset
   Dim Coluna, Linha, Largura_Campo

   VLR_TOT_FAT_N = 0

   grdFaturados.Clear
   grdFaturados.Visible = False
   grdFaturados.Gridlines = flexGridFlat
   grdFaturados.FixedRows = 1
   grdFaturados.FixedCols = 1
   grdFaturados.ScrollBars = flexScrollBarBoth
   grdFaturados.AllowUserResizing = flexResizeColumns

   If TabGrid.State = 1 Then _
      TabGrid.Close

   TabGrid.Open strSQL, CONECTA_RETAGUARDA, adOpenKeyset, adLockOptimistic
   If Not TabGrid.EOF Then
      ' define linhas fixas igual a uma e não usa colunas fixas
      grdFaturados.Rows = 2
      grdFaturados.FixedCols = 0

      ' define o numero de linhas e colunas e cria uma matrix com o total de registros a exibir
      grdFaturados.Rows = 1
      grdFaturados.Cols = TabGrid.Fields.Count

      ReDim largura_coluna(0 To TabGrid.Fields.Count - 1)

      ' exibe os cabeçalhos das colunas
      For Coluna = 0 To TabGrid.Fields.Count - 1
         grdFaturados.TextMatrix(0, Coluna) = Trim(TabGrid.Fields(Coluna).Name)
         largura_coluna(Coluna) = TextWidth(Trim(TabGrid.Fields(Coluna).Name))
      Next Coluna

      ' exibe o valor de cada linha
      Linha = 1

      Do While Not TabGrid.EOF
         grdFaturados.Rows = grdFaturados.Rows + 1

         For Coluna = 0 To TabGrid.Fields.Count - 1
            If Coluna = 1 Then
               grdFaturados.TextMatrix(Linha, Coluna) = Format(TabGrid.Fields(Coluna).Value, strFormatacao3Digitos)
               VLR_TOT_FAT_N = VLR_TOT_FAT_N + TabGrid.Fields(Coluna).Value
               Else: grdFaturados.TextMatrix(Linha, Coluna) = "" & Trim(TabGrid.Fields(Coluna).Value)
            End If

            'If TabGrid.Fields("ST_PEDIDO").Value < 3 Then
            '   grdFaturados.Row = Linha
            '   grdFaturados.Col = Coluna
            '   grdFaturados.CellForeColor = vbRed ' &H4000&   '&H40&
            'End If

            ' verifica o tamanho dos campos
            If Not IsNull(TabGrid.Fields(Coluna).Value) Then _
               Largura_Campo = TextWidth(TabGrid.Fields(Coluna).Value)

            If largura_coluna(Coluna) < Largura_Campo Then _
               largura_coluna(Coluna) = Largura_Campo
         Next Coluna

         TabGrid.MoveNext
         Linha = Linha + 1
      Loop

      If TabGrid.State = 1 Then _
         TabGrid.Close

      'define a largura das colunas do grid
      For Coluna = 0 To grdFaturados.Cols - 1
         grdFaturados.ColWidth(Coluna) = largura_coluna(Coluna) + 240
      Next Coluna

      grdFaturados.ColWidth(0) = 0
      grdFaturados.Refresh

      grdFaturados.BackColor = vbWhite
      grdFaturados.ForeColor = vbBlue
   
'tipovenda
      grdFaturados.ColWidth(0) = 4000
      grdFaturados.ColAlignment(0) = 0

'total
      grdFaturados.ColWidth(1) = 3000
      grdFaturados.ColAlignment(1) = 7

      MOSTRA_TOTAIS Codg_Usu_N

      DoEvents

      grdFaturados.Visible = True
   End If

   cmdFaturados.Caption = "Faturado = " & Format(VLR_TOT_FAT_N, strFormatacao2Digitos)
   cmdFaturados.Refresh

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "SETA_GRID_FECHA_FAT"
End Sub
