VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSConsultaVeiculo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Veículo"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "CONSULTAVEICULO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPLACA 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox cmbAuxTipo 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCHASSI 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtNome 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1320
      Width           =   5295
   End
   Begin VB.TextBox txtANO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtMODELO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox cmbTIPO 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1800
      Width           =   3495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CONSULTAVEICULO.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CONSULTAVEICULO.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CONSULTAVEICULO.frx":077C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CONSULTAVEICULO.frx":0BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CONSULTAVEICULO.frx":1024
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CONSULTAVEICULO.frx":1344
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CONSULTAVEICULO.frx":1798
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
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
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSMask.MaskEdBox txtCGCCPF 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView LISTACHASSI 
      Height          =   3825
      Left            =   0
      TabIndex        =   13
      Top             =   2280
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   6747
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777152
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Placa"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Chassi"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "ANO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "MODELO"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "TIPO"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   15
      Top             =   840
      Width           =   675
   End
   Begin VB.Label lblCpf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   12
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Veículo:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1800
      Width           =   1320
   End
End
Attribute VB_Name = "frmOSConsultaVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call CentralizaJanela(frmOSConsultaVeiculo)

   SETA_GRID_CHASSI
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape: Unload Me
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DBARQAUX.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub cmbTIPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(LISTACHASSI.SelectedItem.Text) Then
         KeyAscii = 0
         SQL3 = LISTACHASSI.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub LISTACHASSI_DblClick()
   If Not IsNull(LISTACHASSI.SelectedItem.Text) Then
      SQL3 = LISTACHASSI.SelectedItem.Text
      Unload Me
   End If
End Sub

Private Sub txtANO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(LISTACHASSI.SelectedItem.Text) Then
         KeyAscii = 0
         SQL3 = LISTACHASSI.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(LISTACHASSI.SelectedItem.Text) Then
         KeyAscii = 0
         SQL3 = LISTACHASSI.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtCHASSI_Change()
   CRITERIO_A = "" & Chr$(39) & txtCHASSI.Text & "*" & Chr(39)
   SETA_GRID_CHASSI
End Sub

Private Sub txtCgccpf_Change()
   SETA_GRID_CHASSI
End Sub

Private Sub txtANO_Change()
   SETA_GRID_CHASSI
End Sub

Private Sub txtCHASSI_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Not IsNull(LISTACHASSI.SelectedItem.Text) Then
         KeyAscii = 0
         SQL3 = LISTACHASSI.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtMODELO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNull(LISTACHASSI.SelectedItem.Text) Then
         KeyAscii = 0
         SQL3 = LISTACHASSI.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> 8 Then
      CRITERIO_A = "" & txtPLACA.Text
      If Len(CRITERIO) = 3 Then
         txtPLACA.Text = CRITERIO_A & "-"
         txtPLACA.SelStart = 4
         txtPLACA.Refresh
      End If
   End If
   If KeyAscii = 13 Then
      If Not IsNull(LISTACHASSI.SelectedItem.Text) Then
         KeyAscii = 0
         SQL3 = LISTACHASSI.SelectedItem.Text
         Unload Me
      End If
   End If
End Sub

Private Sub txtplaca_Change()
   SETA_GRID_CHASSI
End Sub

Private Sub txtmodelo_Change()
   SETA_GRID_CHASSI
End Sub

Private Sub cmbauxtipo_Change()
   SETA_GRID_CHASSI
End Sub

Private Sub SETA_GRID_CHASSI()
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "Aguarde, Pesquisando ..."
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   NUMR_SEQ_N = 1
   NUMR_CONSULTA_N = 0
   HORA_INI = Time
   LISTACHASSI.ListItems.Clear
   SQL = "select * from VEICULO "
   SQL = SQL & "where chassi <> '' "
   If txtCHASSI.Text <> "" Then _
      SQL = SQL & "  and chassi like " & CRITERIO
   txtCGCCPF.PromptInclude = False
   If txtCGCCPF.Text <> "" Then _
      SQL = SQL & " and cgccpf = '" & txtCGCCPF.Text & "'"
   If txtANO.Text <> "" Then _
      SQL = SQL & " and ano = " & txtANO.Text
   If txtMODELO.Text <> "" Then _
      SQL = SQL & " and modelo = " & txtMODELO.Text
   If cmbAuxTipo.Text <> "" Then _
      SQL = SQL & " and tipo = " & cmbAuxTipo.Text
   If txtPLACA.Text <> "" Then _
      SQL = SQL & " and placa like " & Chr$(39) & Replace(txtPLACA.Text, "-", "") & "*" & Chr(39)
   SQL = SQL & " order by ano asc "
   Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
   While Not TabAUX.EOF
      NUMR_CONSULTA_N = NUMR_CONSULTA_N + 1
      Set item = LISTACHASSI.ListItems.Add(, "seq." & TabAUX!placa, TabAUX!placa)
      item.SubItems(1) = TabAUX!chassi

      SQL = "select nome from CLIENTE "
      SQL = SQL & "where cgccpf = '" & TabAUX!CGCCPF & "'"
      Set TabCli = DBARQEMP.OpenRecordset(SQL, 4)
      If Not TabCli.EOF Then _
         item.SubItems(2) = TabCli!NOME
      TabCli.Close

      If Not IsNull(TabAUX!Ano) Then _
         item.SubItems(3) = TabAUX!Ano
      If Not IsNull(TabAUX!Modelo) Then _
         item.SubItems(4) = TabAUX!Modelo
      If Not IsNull(TabAUX!TIPO) Then _
         item.SubItems(5) = TabAUX!TIPO
      TabAUX.MoveNext
   Wend
   TabAUX.Close
   txtCGCCPF.PromptInclude = True
   HORA_FIM = Time

   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = "ESC - Sair"
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (2)
   frmINICIO.BARI.Panels(2).Text = ""
   frmINICIO.BARI.Panels(2).AutoSize = sbrContents

   frmINICIO.BARI.Panels.Add (3)
   frmINICIO.BARI.Panels(3).Text = "Duração da consulta = " & Format((HORA_FIM - HORA_INI), "hh:mm:ss")
   frmINICIO.BARI.Panels(3).AutoSize = sbrContents
         
   frmINICIO.BARI.Panels.Add (4)
   frmINICIO.BARI.Panels(4).Text = "Total de Registros Encontrados = " & NUMR_CONSULTA_N
   frmINICIO.BARI.Panels(4).AutoSize = sbrContents
End Sub
