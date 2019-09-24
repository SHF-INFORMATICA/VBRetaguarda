VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNFeMANUT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manutenção NFe"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NFeMANUT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   15
      TabIndex        =   8
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Atualização"
      TabPicture(0)   =   "NFeMANUT.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtVolume"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtEspecie"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtNota"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPesoBruto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPesoLiquido"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtProtocolo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Atualiza Situação NFe"
      TabPicture(1)   =   "NFeMANUT.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtNFe"
      Tab(1).Control(1)=   "Label1(0)"
      Tab(1).Control(2)=   "Label1(2)"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtProtocolo 
         Height          =   405
         Left            =   1440
         TabIndex        =   18
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtPesoLiquido 
         Height          =   405
         Left            =   4080
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtPesoBruto 
         Height          =   405
         Left            =   4080
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtNota 
         BackColor       =   &H00FFFFC0&
         Height          =   405
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtEspecie 
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtVolume 
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtNFe 
         Height          =   405
         Left            =   -73200
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo:"
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Situação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   5
         Left            =   3240
         TabIndex        =   17
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pesso Liquido:"
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Peso Bruto:"
         Height          =   285
         Left            =   2640
         TabIndex        =   15
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   0
         X2              =   5400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número NFe:"
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Espécie:"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Volume:"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número NFe:"
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Situação:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   2
         Left            =   -74760
         TabIndex        =   10
         Top             =   1200
         Width           =   1320
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
            Key             =   "voltar"
            Object.ToolTipText     =   "Fechar janela"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "limpar"
            Object.ToolTipText     =   "Limpar formulário"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Atualizar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5280
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFeMANUT.frx":5C4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFeMANUT.frx":7072
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFeMANUT.frx":8101
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFeMANUT.frx":9369
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFeMANUT.frx":A474
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFeMANUT.frx":BB71
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "NFeMANUT.frx":CD0B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      Caption         =   "Atualização NFe GLOBAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   525
      Index           =   1
      Left            =   -105
      TabIndex        =   6
      Top             =   720
      Width           =   5580
   End
End
Attribute VB_Name = "frmNFeMANUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then _
        txtNOTA.SetFocus
    If SSTab1.Tab = 1 Then _
        txtNFe.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
        txtProtocolo.Text = ""
         txtNFe.Text = ""
         Label1(2).Caption = "Situação:"
         txtNOTA.Text = ""
         txtVolume.Text = ""
         TxtEspecie.Text = ""
         TxtPesoBruto.Text = ""
         TxtPesoLiquido.Text = ""
      Case "gravar"
         If SSTab1.Tab = 0 Then
            ATUALIZA_NFE_2
         End If
         If SSTab1.Tab = 1 Then
            ATUALIZA_NFE
         End If
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar1_ButtonClick"
End Sub

Private Sub txtNFe_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_NFE
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNFe_KeyPress"
End Sub

Sub PROCURA_NFE()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_GLOBAL

   Label1(2).Caption = "Situação:"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select MFADOC,MFACODSTAT, MFACODMORE, MFACHAVENFE, MFAMOTRESU, MFACODRECI, MFALOTENFE"
   SQL = SQL & " from MFA010"
   SQL = SQL & " where mfadoc = '" & Trim(txtNFe.Text) & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabTemp.EOF Then
      Label1(2).Caption = "Situação: " & Trim(TabTemp.Fields("MFACODMORE").Value) & " - " & _
                                         Trim(TabTemp.Fields("MFAMOTRESU").Value)
      Label1(3).Caption = Trim(TabTemp.Fields("MFACODSTAT").Value)
   End If

   If TabTemp.State = 1 Then _
      TabTemp.Close

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_NFE"
End Sub

Sub ATUALIZA_NFE()
'On Error GoTo ERRO_TRATA

   If Label1(3).Caption = "105" And Trim(txtNFe.Text) <> "" Or Label1(3).Caption = "103" And Trim(txtNFe.Text) <> "" Then
      Msg = "Antes de confirmar essa rotina vc deve verificar se a nota fiscal eletrônica esta validado no site da receita federal (http://www.nfe.fazenda.gov.br/portal/principal.aspx). Confirma Operação ?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         ABRE_BANCO_GLOBAL

            SQL = "update MFA010 set "
            SQL = SQL & " MFACODSTAT = 100, MFACODMORE = 100"
            SQL = SQL & " where mfadoc = '" & Trim(txtNFe.Text) & "'"
            CONECTA_GLOBAL.Execute SQL

         If CONECTA_GLOBAL.State = 1 Then _
            CONECTA_GLOBAL.Close

         MsgBox "Operação realizada com sucesso, feche o aplicativo e emissão NFe e abra novamente."
      End If
      Else
         'If Label1(3).Caption = "106" And Trim(txtNFe.Text) <> "" Then
         If Trim(txtNFe.Text) <> "" Then
            ABRE_BANCO_GLOBAL

               SQL = "update MFA010 set "
               SQL = SQL & " MFACODSTAT = 100, MFACODMORE = 100"
               SQL = SQL & " where mfadoc = '" & Trim(txtNFe.Text) & "'"
               CONECTA_GLOBAL.Execute SQL

            If CONECTA_GLOBAL.State = 1 Then _
               CONECTA_GLOBAL.Close

            MsgBox "Operação realizada com sucesso, feche o aplicativo e emissão NFe e abra novamente."
            'Else: MsgBox "Não permitido para situação : " & Label1(3).Caption
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_NFE"
End Sub
'================================================================================
Private Sub txtNota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
      PROCURA_NFE_2
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNota_KeyPress"
End Sub

Sub PROCURA_NFE_2()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_GLOBAL

   Label1(5).Caption = "Situação:"

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select MFADOC,MFACODSTAT, MFACODMORE, MFACHAVENFE, MFAMOTRESU, MFACODRECI, MFALOTENFE"
   SQL = SQL & " from MFA010"
   SQL = SQL & " where mfadoc = '" & Trim(txtNOTA.Text) & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabTemp.EOF Then _
      Label1(5).Caption = "Situação: OK"
   If TabTemp.State = 1 Then _
      TabTemp.Close

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PROCURA_NFE_2"
End Sub

Sub ATUALIZA_NFE_2()
'On Error GoTo ERRO_TRATA

   If Trim(txtNOTA.Text) <> "" Then
      Msg = "Confirmar Operação ?"
      PERGUNTA Msg, vbYesNo + 32, "Atenção !!!", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         ABRE_BANCO_GLOBAL

            SQL = "update MFA010 set "
            SQL = SQL & " mfavolume4 = '" & Trim(txtVolume.Text) & "'"
            SQL = SQL & ", MFAESPECIE = '" & Trim(TxtEspecie.Text) & "'"
            
            SQL = SQL & ", MFAPBRUTO = '" & Trim(TxtPesoBruto.Text) & "'"
            SQL = SQL & ", MFAPLIQUI = '" & Trim(TxtPesoLiquido.Text) & "'"

            If Trim(txtProtocolo.Text) <> "" Then _
                SQL = SQL & ", MFACODPROT = '" & Trim(txtProtocolo.Text) & "'"

            SQL = SQL & " where mfadoc = '" & Trim(txtNOTA.Text) & "'"
            CONECTA_GLOBAL.Execute SQL

         If CONECTA_GLOBAL.State = 1 Then _
            CONECTA_GLOBAL.Close

         MsgBox "Operação realizada com sucesso, feche o aplicativo e emissão NFe e abra novamente."
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ATUALIZA_NFE_2"
End Sub
