VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCCe 
   Caption         =   "CC-e | Carta de correção eletrônica"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CCe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkImp 
      Caption         =   "Impressora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6840
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtNumrNota 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin TabDlg.SSTab tab_CCe 
      Height          =   3645
      Left            =   0
      TabIndex        =   5
      Top             =   1500
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6429
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CC-e"
      TabPicture(0)   =   "CCe.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtCorrecao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "CCe.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grade"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid Grade 
         Height          =   3255
         Left            =   -74950
         TabIndex        =   8
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5741
         _Version        =   393216
      End
      Begin VB.TextBox txtCorrecao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   50
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   9255
      End
   End
   Begin VB.TextBox txtChave 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   6250
   End
   Begin MSMask.MaskEdBox txtDtEmis 
      Height          =   330
      Left            =   1695
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1270
      ButtonWidth     =   2725
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
            Caption         =   "&Imprimir"
            Key             =   "print"
            Object.ToolTipText     =   "Impressão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerar"
            Key             =   "gravar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Atualizar"
            Key             =   "at"
            ImageIndex      =   4
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
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CCe.frx":5C4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CCe.frx":6DE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CCe.frx":7E73
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CCe.frx":8E28
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CCe.frx":9F33
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CCe.frx":BF15
               Key             =   ""
            EndProperty
         EndProperty
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
      DesignWidth     =   9375
      DesignHeight    =   5130
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   0
      X2              =   9360
      Y1              =   1450
      Y2              =   1450
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   9360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lbl_DadosNFe 
      AutoSize        =   -1  'True
      Caption         =   "Número NF-e          Data Emissão      Chave de acesso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   45
      TabIndex        =   4
      Top             =   840
      Width           =   4470
   End
End
Attribute VB_Name = "frmCCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SET AQUI VER PORQUE A PORRA DO PROTOCOLO DÁ PAU
Option Explicit
   Public Numr_Nota_N   As Long
   Const strRegistrado = "EVENTO REGISTRADO E VINCULADO"

Private Sub Form_Load()
'On Error GoTo ERRO_TRATA

   ABRE_BANCO_GLOBAL

   If EXISTE_OBJ_BANCO("GLOBAL", "TB_CARTACORRECAO_NFE", "") = False Then
      SQL = " CREATE TABLE [dbo].[TB_CARTACORRECAO_NFE]("
      SQL = SQL & " [EMPRESA_ID] [char](2) NOT NULL,"          'codigo da empresa
      SQL = SQL & " [NUMR_NOTA] [varchar](9) NOT NULL,"        'numero da chave da nota ou voce pode usar o numero da nota fiscal de saida tem que ter constraint com seu arquivo de nota
      SQL = SQL & " [CODG_CORRECAO] [bigint] NOT NULL,"        'codigo da carta de correcao, sendo que este peercente a chave empresaID+Numeronfe+codg_correcao (pode ter varios para cada numero de nota)
      SQL = SQL & " [REGISTRO_CORRECAO] [datetime] NOT NULL,"  'aqui voce vai grava a data e hora daocorrencia
      SQL = SQL & " [DADOS_CORRECAO] [text] NOT NULL,"
      SQL = SQL & " [NUMERO_PROTOCOLO] [varchar](50) NULL,"
      SQL = SQL & " [COD_STATUS] [int] NULL,"
      SQL = SQL & " [DATA_EVENTO] [datetime] NULL,"
      SQL = SQL & " [CCEMOTRESU] [varchar](600) NULL"
      SQL = SQL & " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
      CONECTA_GLOBAL.Execute SQL
   End If

   txtCorrecao.Text = "Informe aqui as correções que deseje ajustar na NF-e, no mínimo com 15 dígitos."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ERRO_TRATA

   Numr_Nota_N = 0

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Form_Unload"
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ERRO_TRATA

   Select Case Button.key
      Case "at"
         ROTINA_VACA
      Case "gravar"
         Call GERA_CCe
         PREENCHE_GRID
         LIMPA_TUDO
      Case "limpar"
         LIMPA_TUDO
      Case "print"
         FORMULA_REL = "{TB_CARTACORRECAO_NFE_old.empresa_id} = " & EMPRESA_ID_N
         FORMULA_REL = FORMULA_REL & " and {TB_CARTACORRECAO_NFE_old.NUMR_NOTA} = '" & Trim(txtNumrNota.Text) & "'"

         ESCOLHE_IMPRESSORA "global"

         Nome_Relatorio = "rel_CCe.rpt"
         frmRELATORIO10.Show 1
      Case "voltar"
         Unload Me
   End Select

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "Toolbar_ButtonClick"
End Sub

Private Sub txtNumrNota_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0

      If Trim(txtNumrNota.Text) <> "" Then _
         If IsNumeric(txtNumrNota.Text) Then _
            Numr_Nota_N = txtNumrNota.Text

      SendKeys "{tab}"
      Else
         If KeyAscii = 8 Then
            Else: If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
         End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNumrNota_KeyPress"
End Sub

Private Sub txtNumrNota_LostFocus()
'On Error GoTo ERRO_TRATA

   If Trim(txtNumrNota.Text) <> "" Then
      If IsNumeric(txtNumrNota.Text) Then
         Numr_Nota_N = txtNumrNota.Text

         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from MFA010 "
         SQL = SQL & " where MFADOC = '" & Trim(Numr_Nota_N) & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

         TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabTemp.EOF Then
            txtDtEmis.PromptInclude = False
               txtDtEmis.Text = "" & Trim(TabTemp.Fields("mfaemissao").Value)
            txtDtEmis.PromptInclude = True

            txtChave.Text = "" & Trim(TabTemp.Fields("mfachavenfe").Value)
            Else: MsgBox "Nota fiscal não encontrada."
         End If
         If TabTemp.State = 1 Then _
            TabTemp.Close

         PREENCHE_GRID
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNumrNota_LostFocus"
End Sub

Private Sub txtCorrecao_KeyPress(KeyAscii As Integer)
'On Error GoTo ERRO_TRATA

   If KeyAscii = 13 Then
      KeyAscii = 0
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "txtNumrNota_KeyPress"
End Sub

Private Function GERA_CCe()
'On Error GoTo ERRO_TRATA

   If Trim(txtNumrNota.Text) <> "" Then _
      If IsNumeric(txtNumrNota.Text) Then _
         Numr_Nota_N = txtNumrNota.Text

   If Trim(txtNumrNota.Text) = "" Then
      MsgBox "Informe número de nota."
      Exit Function
   End If

   If Len(Trim(txtCorrecao.Text)) >= 15 Then
      NUMR_SEQ_N = 1

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select max(codg_correcao) from TB_CARTACORRECAO_NFE "
      SQL = SQL & " WHERE numr_nota = " & Numr_Nota_N
      TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabTemp.EOF Then _
         If Not IsNull(TabTemp.Fields(0).Value) Then _
            NUMR_SEQ_N = TabTemp.Fields(0).Value + 1

      If TabTemp.State = 1 Then _
         TabTemp.Close

      If NUMR_SEQ_N >= 21 Then
         MsgBox "Não foi possível gerar a carta de correção eletrônica!" & vbLf & _
         "Motivo:Excedido o limite de 20 correções para esta NF-e!", 48
         Exit Function
      End If

      SQL = "insert into TB_CARTACORRECAO_NFE "
      SQL = SQL & "("
         SQL = SQL & " EMPRESA_ID,NUMR_NOTA,CODG_CORRECAO,REGISTRO_CORRECAO,DADOS_CORRECAO"
         'SQL = SQL & " NUMERO_PROTOCOLO,COD_STATUS,DATA_EVENTO,CCEMOTRESU"
      SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & "0" & EMPRESA_ID_N                     'EMPRESA_ID
         SQL = SQL & ",'" & Numr_Nota_N & "'"               'NUMR_NOTA
         SQL = SQL & "," & NUMR_SEQ_N                       'CODG_CORRECAO
         SQL = SQL & ",'" & DMA(Date) & "'"                 'REGISTRO_CORRECAO
         SQL = SQL & ",'" & Trim(Replace(txtCorrecao.Text, ",", ".")) & "'"  'DADOS_CORRECAO

         'NUMERO_PROTOCOLO
         'COD_STATUS
         'DATA_EVENTO
         'CCEMOTRESU

      SQL = SQL & ")"

      CONECTA_GLOBAL.Execute SQL

      MsgBox "Carta de correção registrada com sucesso!"
      txtCorrecao.Text = "Informe aqui as correções que deseje ajustar na NF-e, no mínimo com 15 dígitos."
      
      Else: MsgBox "Informe a correção com no mínimo com 15 dígitos!", 48
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "GERA_CCe"
End Function

Private Sub FORMA_GRID()
'On Error GoTo ERRO_TRATA

   Grade.AllowUserResizing = flexResizeColumns
   Grade.Clear
   Grade.Rows = 2
   Grade.Cols = 8
   
   Grade.TextMatrix(0, 0) = "Cod.Correção"
   Grade.ColWidth(0) = 1200
   Grade.ColAlignment(0) = flexAlignLeftCenter
   
   Grade.TextMatrix(0, 1) = "Data"
   Grade.ColWidth(1) = 1260
   Grade.ColAlignment(1) = flexAlignCenterCenter
   
   Grade.TextMatrix(0, 2) = "Horas"
   Grade.ColWidth(2) = 1000
   Grade.ColAlignment(2) = flexAlignLeftCenter
   
   Grade.TextMatrix(0, 3) = "Ajustes"
   Grade.ColWidth(3) = 10000
   Grade.ColAlignment(3) = flexAlignLeftCenter

   Grade.TextMatrix(0, 4) = "Protocolo"
   Grade.ColWidth(4) = 10000
   Grade.ColAlignment(4) = flexAlignLeftCenter

   Grade.TextMatrix(0, 5) = "Código Sefaz"
   Grade.ColWidth(5) = 10000
   Grade.ColAlignment(5) = flexAlignLeftCenter
         
   Grade.TextMatrix(0, 6) = "Data Evento"
   Grade.ColWidth(6) = 10000
   Grade.ColAlignment(6) = flexAlignLeftCenter

   Grade.TextMatrix(0, 7) = "Motivo Rejeição"
   Grade.ColWidth(7) = 10000
   Grade.ColAlignment(7) = flexAlignLeftCenter

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FORMA_GRID"
End Sub

Private Sub PREENCHE_GRID()
'On Error GoTo ERRO_TRATA

   Dim i As Long

   Call FORMA_GRID
   i = 1

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TB_CARTACORRECAO_NFE "
   SQL = SQL & " WHERE numr_nota = " & Numr_Nota_N
   SQL = SQL & " ORDER BY codg_correcao ASC "
   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   With TabTemp
      While Not .EOF
         Grade.TextMatrix(i, 0) = !CODG_CORRECAO
         Grade.TextMatrix(i, 1) = Format(Mid(!REGISTRO_CORRECAO, 1, 10), "dd/MM/yyyy")
         Grade.TextMatrix(i, 2) = Mid(!REGISTRO_CORRECAO, 12, 8)
         Grade.TextMatrix(i, 3) = !DADOS_CORRECAO

         'NUMERO_PROTOCOLO
         'COD_STATUS
         'DATA_EVENTO
         'CCEMOTRESU

         i = i + 1
         Grade.Rows = Grade.Rows + 1

         .MoveNext
      Wend
   End With
   If TabTemp.State = 1 Then _
      TabTemp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PREENCHE_GRID"
End Sub

Sub LIMPA_TUDO()
'On Error GoTo ERRO_TRATA

   txtCorrecao.Text = ""
   txtNumrNota.Text = ""
   txtDtEmis.PromptInclude = False
      txtDtEmis.Text = ""
   txtDtEmis.PromptInclude = True
   txtChave.Text = ""
   txtNumrNota.SetFocus

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_TUDO"
End Sub

Sub ROTINA_VACA()
'On Error GoTo ERRO_TRATA

   Dim NUMERO_PROTOCOLO, CCEMOTRESU
   Dim DATA_EVENTO As String

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from TB_CARTACORRECAO_NFE "
   SQL = SQL & " ORDER BY codg_correcao ASC "
   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   With TabTemp
      While Not .EOF

         NUMERO_PROTOCOLO = ""
         DATA_EVENTO = ""
         CCEMOTRESU = ""

         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         SQL = "select * from [TB_CARTACORRECAO_NFE_old] "
         SQL = SQL & " where numr_nota = " & TabTemp.Fields("numr_nota").Value
         TabConsulta.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If TabConsulta.EOF Then
            SQL3 = Replace(TabTemp.Fields("DADOS_CORRECAO").Value, "'", ";")
            SQL3 = Replace(SQL3, ",", ";")

            If IsNull(TabTemp.Fields("NUMERO_PROTOCOLO").Value) Then
               NUMERO_PROTOCOLO = "null"
               Else: NUMERO_PROTOCOLO = Trim(TabTemp.Fields("NUMERO_PROTOCOLO").Value)
            End If
            If IsNull(TabTemp.Fields("DATA_EVENTO").Value) Then
               DATA_EVENTO = "null"
               Else: DATA_EVENTO = DMA(TabTemp.Fields("DATA_EVENTO").Value)
            End If
            If IsNull(TabTemp.Fields("CCEMOTRESU").Value) Then
               CCEMOTRESU = "null"
               Else: CCEMOTRESU = Trim(TabTemp.Fields("CCEMOTRESU").Value)
            End If

            SQL = "insert into TB_CARTACORRECAO_NFE_old "
               SQL = SQL & "(EMPRESA_ID,NUMR_NOTA,CODG_CORRECAO,REGISTRO_CORRECAO,DADOS_CORRECAO,"
               SQL = SQL & "NUMERO_PROTOCOLO,COD_STATUS,DATA_EVENTO,CCEMOTRESU)"
            SQL = SQL & " values("
               SQL = SQL & TabTemp.Fields("empresa_id").Value              'EMPRESA_ID
               SQL = SQL & "," & TabTemp.Fields("NUMR_NOTA").Value         'NUMR_NOTA
               SQL = SQL & "," & TabTemp.Fields("CODG_CORRECAO").Value     'CODG_CORRECAO
               SQL = SQL & "," & TabTemp.Fields("REGISTRO_CORRECAO").Value 'REGISTRO_CORRECAO,
               SQL = SQL & ",'" & SQL3 & "'"                               'DADOS_CORRECAO,
               SQL = SQL & ",'" & Trim(NUMERO_PROTOCOLO) & "'"             'NUMERO_PROTOCOLO
               SQL = SQL & "," & TabTemp.Fields("COD_STATUS").Value        'COD_STATUS
               SQL = SQL & ",'" & DMA(DATA_EVENTO) & "'"                     'DATA_EVENTO
               SQL = SQL & ",'" & Trim(CCEMOTRESU) & "'"                   'CCEMOTRESU
            SQL = SQL & ")"
'MsgBox SQL
            CONECTA_GLOBAL.Execute SQL
         End If
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         If Not IsNull(TabTemp.Fields("NUMERO_PROTOCOLO").Value) Then
            SQL = "update TB_CARTACORRECAO_NFE set "
            SQL = SQL & " numero_protocolo = '" & Trim(TabTemp.Fields("NUMERO_PROTOCOLO").Value) & "'"
            SQL = SQL & " where numr_nota = " & Trim(TabTemp.Fields("numr_nota").Value)
            CONECTA_GLOBAL.Execute SQL
         End If

         .MoveNext
      Wend
   End With
   If TabTemp.State = 1 Then _
      TabTemp.Close

MsgBox "Ok, atualizado pronto para impressão."

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "ROTINA_VACA"
End Sub
