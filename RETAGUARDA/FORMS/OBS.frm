VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOBS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observações"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OBS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOBS 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   7815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
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
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Gravar"
            Key             =   "gravar"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   10080
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":11A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":2235
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":31EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":42F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":544B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":589D
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":7714
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":8DCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "OBS.frx":ADAC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Impressora"
         Height          =   240
         Left            =   10200
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*Vendedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmOBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If CHAMADA_A = "TERMOGARANTIA" Then
      Me.Caption = "Termo Garantia O.S. nº: " & OS_ID_N
      
      SQL = "Garantia de Serviços 30 Dias "
      SQL = SQL & vbCrLf
      SQL = SQL & vbCrLf
      SQL = SQL & "Garantia de Peças 90 dias sem impurezas no óleo."
      txtOBS.Text = SQL

      If TabDESCR.State = 1 Then _
         TabDESCR.Close
      SQL = "select * from DESCR "
      SQL = SQL & " where tipo = 'A3' "
      SQL = SQL & "order by codigo "
      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         txtOBS.Text = Trim(TabDESCR!DESCRICAO)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      SQL = "select OSTERMOOBS from OSTERMO "
      SQL = SQL & " where os_id = " & OS_ID_N
      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCabeca.EOF Then _
         frmOBS.txtOBS.Text = "" & TabCabeca.Fields("OSTERMOOBS").Value
      If TabCabeca.State = 1 Then _
         TabCabeca.Close
      Else
         If CHAMADA_A = "OBS" Then
            Me.Caption = "Observações O.S. nº: " & OS_ID_N
            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            SQL = "select obs from OSOBS "
            SQL = SQL & " where os_id = " & OS_ID_N
            TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCabeca.EOF Then _
               frmOBS.txtOBS.Text = "" & TabCabeca.Fields("obs").Value
            If TabCabeca.State = 1 Then _
               TabCabeca.Close
         End If
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape
         Unload Me
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   OBS_A = Trim(txtOBS.Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         txtOBS.Text = ""
      Case "gravar"
         NUMR_ID_N = 0

         If CHAMADA_A = "TERMOGARANTIA" Then
            NUMR_ID_N = 0

            If TabCabeca.State = 1 Then _
               TabCabeca.Close
            SQL = "select * from OSTERMO "
            SQL = SQL & " where os_id = " & OS_ID_N
            TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabCabeca.EOF Then
               Acao_N = 1
               Else
                  Acao_N = 2
                  NUMR_ID_N = 0 & TabCabeca.Fields("OSTERMO_ID").Value
            End If
            If TabCabeca.State = 1 Then _
               TabCabeca.Close

            spOSTERMO Acao_N, NUMR_ID_N, OS_ID_N, Trim(OBS_A)

            Else
               If CHAMADA_A = "OBS" Then
                  NUMR_ID_N = 0

                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close
                  SQL = "select * from OSOBS "
                  SQL = SQL & " where os_id = " & OS_ID_N
                  TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabCabeca.EOF Then
                     Acao_N = 1
                     Else
                        Acao_N = 2
                        NUMR_ID_N = 0 & TabCabeca.Fields("OSOBS_ID").Value
                  End If
                  If TabCabeca.State = 1 Then _
                     TabCabeca.Close

                  spOSOBS Acao_N, NUMR_ID_N, OS_ID_N, Trim(OBS_A)
               End If
         End If
         Unload Me
   End Select
End Sub

Private Sub txtOBS_Change()
   OBS_A = txtOBS.Text
End Sub
