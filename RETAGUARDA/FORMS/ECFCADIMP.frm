VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmECFCADIMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros Impressora Fiscal"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ECFCADIMP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescricao 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton cmdCaixa 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   2340
      Picture         =   "ECFCADIMP.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   405
   End
   Begin VB.TextBox txtCaixa 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdReinicio 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5640
      Picture         =   "ECFCADIMP.frx":6614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   405
   End
   Begin VB.TextBox txtReinicio 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton cmdAliquota 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5640
      Picture         =   "ECFCADIMP.frx":7016
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   405
   End
   Begin VB.TextBox txtAliquotas 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2640
      Width           =   4335
   End
   Begin VB.CommandButton cmdSerie 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5640
      Picture         =   "ECFCADIMP.frx":7A18
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   405
   End
   Begin VB.TextBox txtSerie 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   1270
      ButtonWidth     =   2328
      ButtonHeight    =   1111
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Voltar"
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
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "gravar"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   36
         ImageHeight     =   36
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":841A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":95B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":A643
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":B5F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":C703
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":D859
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":DCAB
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ECFCADIMP.frx":FB22
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid grdECF 
      Bindings        =   "ECFCADIMP.frx":111D8
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoECF 
      Height          =   330
      Left            =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Grid Cabeça"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nº Caixa:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Contador Reinício:"
      Height          =   240
      Index           =   3
      Left            =   165
      TabIndex        =   14
      Top             =   3240
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Aliquotas:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   7800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Parametros Impressora Fiscal"
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
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   6090
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Série ECF:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmECFCADIMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tnn – Tributado (sujeito ao ICMS)
'ISnn – Tributado (sujeito ao ISS)
'F - Substituição Tributária
'i -Isenção
'N - Não incidência;

'Comandos de Cumpom Fiscal
'Abertura de cupom fiscal [00]
'Aumentando a Descrição do Item [6252]
'Acréscimo/Desconto em item posterior [93]
'Cancelamento de Acréscimo/Desconto em item posterior [114]
'Cancelamento de Item anterior [13]
'Cancelamento de Item Genérico [31]
'Cancelamento de Cupom [14]
'Inicia Fechamento de Cupom com Forma de Pgto [32]
'Inicia Fechamento de Cupom sem Forma de Pgto [103]
'Acréscimo/Desconto em subtotal [104]
'Cancelamento de Acréscimo/Desconto em subtotal [105]
'Totaliza o Cupom Fiscal [106]
'Efetua forma de pagamento [72]
'Efetua forma de pagamento com parcelamento [90]
'Termina Fechamento [34]
'Cupom Adicional [85]
'Estorno da Forma de Pagamento [74]

Option Explicit
   Dim NumeroCaixa           As String

Private Sub Form_Load()
   LIMPA_TELA
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "voltar"
         Unload Me
      Case "limpar"
         LIMPA_TELA
      Case "gravar"
         GRAVA_ECF
   End Select
End Sub

Private Sub cmdCaixa_Click()
'On Error GoTo ERRO_TRATA

   NumeroCaixa = Space(4)

   RETORNO_ECF = Bematech_FI_NumeroCaixa(NumeroCaixa)

   txtCaixa.Text = NUMERO_CAIXA_CPU

   If IsNumeric(NUMERO_CAIXA_CPU) Then
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      SQL = "select * from IMPRESSORA "
      SQL = SQL & " where numr_caixa = " & NUMERO_CAIXA_CPU
      TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCAIXA.EOF Then
         txtDescricao.Text = "" & Trim(TabCAIXA.Fields("DESCRICAO").Value)
         txtSerie.Text = "" & Trim(TabCAIXA.Fields("NUMR_SERIE_IMP").Value)
         txtReinicio.Text = "" & Trim(TabCAIXA.Fields("CONTA_REINICIO").Value)
      End If

      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close
   End If
   txtDescricao.Text = "" & "CAIXA" & Int(NUMERO_CAIXA_CPU)

   Call cmdAliquota_Click

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdCaixa_Click"
End Sub

Private Sub cmdSerie_Click()
'On Error GoTo ERRO_TRATA

'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

   Dim NumeroSerie   As String
   Dim LocalRetorno  As String

   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      NumeroSerie = Space(1)
      Else: NumeroSerie = Space(20)
   End If

   'RETORNO_ECF = Bematech_FI_NumeroSerie(NumeroSerie)
   RETORNO_ECF = Bematech_FI_NumeroSerieMFD(NumeroSerie)
   Call VerificaRetornoImpressora("Número de Série: ", NumeroSerie, "Informações da Impressora")
   NUMERO_SERIE_ECF = CStr(NumeroSerie)

NUMERO_SERIE_ECF = CaracteresValidos(NUMERO_SERIE_ECF)

'MsgBox Len(NUMERO_SERIE_ECF)

If Len(NUMERO_SERIE_ECF) < 20 Then _
   NUMERO_SERIE_ECF = NUMERO_SERIE_ECF & NUMERO_CAIXA_CPU

   txtSerie.Text = NUMERO_SERIE_ECF

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdSerie_Click"
End Sub

Private Sub cmdAliquota_Click()
'On Error GoTo ERRO_TRATA
'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

   Dim Aliquotas As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
       Aliquotas = Space(1)
      Else: Aliquotas = Space(79)
   End If

   RETORNO_ECF = Bematech_FI_RetornoAliquotas(Aliquotas)
   Call VerificaRetornoImpressora("Alíquotas Cadastradas: ", Aliquotas, "Informações da Impressora")

   txtAliquotas.Text = "" & Trim(Aliquotas)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdAliquota_Click"
End Sub

Private Sub cmdReinicio_Click()
'On Error GoTo ERRO_TRATA

   Dim NumeroIntervencao As String, CONTA_REINICIO As String
   Dim LocalRetorno As String
   If (LocalRetorno = "1") Then 'Grava retorno em arquivo
      NumeroIntervencao = Space(1)
      Else: NumeroIntervencao = Space(4)
   End If

   RETORNO_ECF = Bematech_FI_NumeroIntervencoes(NumeroIntervencao)

   If Trim(NumeroIntervencao) <> "" Then _
      CONTA_REINICIO = NumeroIntervencao

   txtReinicio.Text = "" & CONTA_REINICIO

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "cmdReinicio_Click"
End Sub

Sub LIMPA_TELA()
   txtCaixa.Text = ""
   txtDescricao.Text = ""
   txtSerie.Text = ""
   txtAliquotas.Text = ""
   txtReinicio.Text = ""
   SETA_GRID
End Sub

Sub GRAVA_ECF()
'On Error GoTo ERRO_TRATA

   If Trim(txtCaixa.Text) <> "" And _
      Trim(txtDescricao.Text) <> "" And _
      Trim(txtSerie.Text) <> "" And _
      Trim(txtAliquotas.Text) <> "" And _
      Trim(txtReinicio.Text) Then

      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      SQL = "select * from IMPRESSORA "
      SQL = SQL & " where numr_caixa = " & NUMERO_CAIXA_CPU
      SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and NUMR_SERIE_IMP = '" & Trim(txtSerie.Text) & "'"
      TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCAIXA.EOF Then
         IMPRESSORA_ID_N = TabCAIXA.Fields("IMPRESSORA_ID").Value

         SQL = "update IMPRESSORA set"

         SQL = SQL & " NUMR_CAIXA = " & Trim(NUMERO_CAIXA_CPU)
         SQL = SQL & ", CONTA_REINICIO = " & Trim(txtReinicio.Text)
         SQL = SQL & ", DESCRICAO = '" & Trim(txtDescricao.Text) & "'"
         SQL = SQL & ", NUMR_SERIE_IMP = '" & Trim(txtSerie.Text) & "'"

         SQL = SQL & " where numr_caixa = " & NUMERO_CAIXA_CPU
         SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and NUMR_SERIE_IMP = '" & Trim(txtSerie.Text) & "'"
         Else
            IMPRESSORA_ID_N = MAX_ID("IMPRESSORA_ID", "IMPRESSORA", "", "", "", "")
   
            SQL = "insert into IMPRESSORA "
            SQL = SQL & "(IMPRESSORA_ID,EMPRESA_ID,NUMR_CAIXA,CONTA_REINICIO,NUMR_SERIE_IMP,DESCRICAO)"
            SQL = SQL & " values("
               SQL = SQL & IMPRESSORA_ID_N                        'IMPRESSORA_ID
               SQL = SQL & "," & EMPRESA_ID_N                     'EMPRESA_ID
               SQL = SQL & "," & Trim(NUMERO_CAIXA_CPU)           'NUMR_CAIXA
               SQL = SQL & "," & Trim(txtReinicio.Text)           'CONTA_REINICIO
               SQL = SQL & ",'" & Trim(txtSerie.Text) & "'"       'NUMR_SERIE_IMP
               SQL = SQL & ",'" & Trim(txtDescricao.Text) & "'"   'DESCRICAO
            SQL = SQL & ")"
      End If
   
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close
   
      CONECTA_RETAGUARDA.Execute SQL
   '-=========
   
      SQL = "delete from INDICE"
      SQL = SQL & " where impressora_id = " & IMPRESSORA_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   
      Dim a(1 To 15)    As String
      Dim Aliquota_N    As Long
      Dim INDICE_ID_N   As Integer
   
      INDICE_ID_N = 1
   
      ParseToArrayVIRGULA txtAliquotas.Text, a()
   
      DoEvents

      While INDICE_ID_N < 15
         If Trim(a(INDICE_ID_N)) <> "" Then
            If IsNumeric(a(INDICE_ID_N)) Then
               Aliquota_N = Left(a(INDICE_ID_N), 2)
               If Aliquota_N > 0 Then
                  If TabCAIXA.State = 1 Then _
                     TabCAIXA.Close

                  SQL = "select * from INDICE"
                  SQL = SQL & " where aliquota = " & Aliquota_N
                  SQL = SQL & " and impressora_id = " & IMPRESSORA_ID_N
                  TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabCAIXA.EOF Then
                     SQL = "insert into INDICE "
                     SQL = SQL & " values ("
                        SQL = SQL & INDICE_ID_N             'INDICE_ID
                        SQL = SQL & "," & IMPRESSORA_ID_N   'IMPRESSORA_ID
                        SQL = SQL & "," & Aliquota_N        'ALIQUOTA
                     SQL = SQL & ")"
                     Else
                        SQL = "update INDICE set "
                        SQL = SQL & " INDICE_ID = " & INDICE_ID_N          'INDICE_ID
                        SQL = SQL & ", IMPRESSORA_ID = " & IMPRESSORA_ID_N 'IMPRESSORA_ID
                        SQL = SQL & ", ALIQUOTA = " & Aliquota_N           'ALIQUOTA
                        SQL = SQL & " where aliquota = " & Aliquota_N
                        SQL = SQL & " and impressora_id = " & IMPRESSORA_ID_N
                  End If

                  If TabCAIXA.State = 1 Then _
                     TabCAIXA.Close

                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
         End If
         INDICE_ID_N = INDICE_ID_N + 1
      Wend

      LIMPA_TELA
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "GRAVA_ECF"
End Sub

Sub SETA_GRID()
'On Error GoTo ERRO_TRATA

   SQL = "select impressora_id,numr_caixa,CONTA_REINICIO,descricao,numr_serie_imp from IMPRESSORA"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N

   adoECF.ConnectionString = AUTENTICA_GRID
   adoECF.CommandType = adCmdText

   adoECF.RecordSource = SQL
   adoECF.Enabled = True
   adoECF.Refresh

   grdECF.Columns(0).DataField = "Impressora_ID"
   grdECF.Columns(0).Caption = "Impressora_ID"
   grdECF.Columns(0).Width = 800

   grdECF.Columns(1).DataField = "NUMR_CAIXA"
   grdECF.Columns(1).Caption = "CAIXA"
   grdECF.Columns(1).Width = 800

   grdECF.Columns(2).DataField = "CONTA_REINICIO"
   grdECF.Columns(2).Caption = "ContadorReinício"
   grdECF.Columns(2).Width = 800

   grdECF.Columns(3).DataField = "DESCRICAO"
   grdECF.Columns(3).Caption = "Descrição"
   grdECF.Columns(3).Width = 1200

   grdECF.Columns(4).DataField = "NUMR_SERIE_IMP"
   grdECF.Columns(4).Caption = "NºSérie"
   grdECF.Columns(4).Width = 2000

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.description, Me.name, "SETA_GRID"
End Sub
