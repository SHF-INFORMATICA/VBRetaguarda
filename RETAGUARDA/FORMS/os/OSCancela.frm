VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOSCancela 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelamento de Ordem de Serviço"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "OSCancela.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPLACA 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   9600
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtOR 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1680
      Width           =   5415
   End
   Begin VB.TextBox txtForma 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   6240
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin MSMask.MaskEdBox DTEMIS 
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox DTCANCELA 
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CGCCPF 
      Height          =   360
      Left            =   4200
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   120
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
            Picture         =   "OSCancela.frx":5C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSCancela.frx":6066
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSCancela.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSCancela.frx":67D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSCancela.frx":6C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSCancela.frx":6F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OSCancela.frx":739E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1164
      ButtonWidth     =   1614
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
            Caption         =   "Confirmar"
            Key             =   "gravar"
            Object.ToolTipText     =   "Efetivação da comissão"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "consultar"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView LISTASERVIÇO 
      Height          =   3705
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   6535
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
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Códg."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tarefa"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Descont."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Mecanico"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LISTAPEÇA 
      Height          =   3705
      Left            =   6000
      TabIndex        =   8
      Top             =   2160
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   6535
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
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Códg."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qtd."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Desconto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Referência"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PLACA:"
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
      Index           =   4
      Left            =   8760
      TabIndex        =   15
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nº Ordem de Serviço:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATA EMISSÃO:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATA CANCELA:"
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
      Index           =   2
      Left            =   2160
      TabIndex        =   12
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENTE:"
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
      Index           =   3
      Left            =   4200
      TabIndex        =   11
      Top             =   1440
      Width           =   1020
   End
End
Attribute VB_Name = "frmOSCancela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call CentralizaJanela2(frmOSCancela)
   frmOSCancela.Top = 1500
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.key
      Case "consultar"
         frmOSCONSULTA.Show 1
         If IsNumeric(CRITERIO) Then
            txtOR.Text = CRITERIO_A
            CRITERIO_A = "" & ""
         End If
      Case "gravar"
         If Not IsNumeric(txtOR.Text) Then _
            Exit Sub
         SQL = "select * from PEDIDO "
         SQL = SQL & " where numr_req = " & txtOR.Text
         Set TabCABECA = DBARQEMP.OpenRecordset(SQL)
         If Not TabCABECA.EOF Then
            If TabCABECA!STATUS = 9 Then
               TabCABECA.Close
               MsgBox "Esse registro já foi cancelado !!!"
               txtOR.SetFocus
               Exit Sub
               Else
                  SQL = "select * from CABECAOS "
                  SQL = SQL & " where pedido_id =  " & PEDIDO_ID_N
                  TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If Not TabTemp.EOF Then
                     If Not IsNull(TabTemp!STATUS) Then
                        If TabTemp!STATUS = "A" Then _
                           txtStatus.Text = "Aberta"
                        If TabTemp!STATUS = "B" Then
                           txtStatus.Text = "BAIXADA"
                           MsgBox "Não é permitido cancelar Ordem de Serviço já cancelada."
                           Unload Me
                        End If
                        If TabTemp!STATUS = "C" Then _
                           txtStatus.Text = "CANCELADA"
                        If TabTemp!STATUS = "D" Then _
                           txtStatus.Text = "NEGOCIAÇÂO"
                         If TabTemp!STATUS = "E" Then _
                           txtStatus.Text = "EXECUSÃO"
                        If TabTemp!STATUS = "F" Then
                           txtStatus.Text = "FECHADA"
                           MsgBox "Não é permitido cancelar Ordem de Serviço fechada."
                           Unload Me
                        End If
                     End If
                  End If

                  'SE NÃO EXISTE NENHUMA REFERENCIA PARA O REGISTRO ELE É DELETADO MESMO
                  'SENÃO É CANCELADO
                  SQL = "select numr_req from PEDIDO "
                  SQL = SQL & " where numr_req not in (" & "select numr_req from CUPOM " & ")"
                  SQL = SQL & "  and numr_req not in (" & "select numr_req from NOTATEMP " & ")"
                  SQL = SQL & "  and numr_req not in (" & "select numr_doc from LANCAMENTO " & ")"
                  SQL = SQL & "  and numr_req = " & TabCABECA!NUMR_REQ
                  Set TabNOTA = DBARQEMP.OpenRecordset(SQL)
                  If Not TabNOTA.EOF Then
                     Msg = "Deseja excluir esse registro? Ordem de Serviço = " & TabCABECA!NUMR_REQ
                     'PERGUNTA
                     If RESPOSTA = vbYes Then
                        TabCABECA.Delete
                        TabCABECA.Close
                        TabNOTA.Close
                        MsgBox "Ordem de Serviço excluida com sucesso."
                        Exit Sub
                     End If
                  End If
                  TabNOTA.Close

                  Msg = "Confirma cancelamento ?"
                  Style = vbYesNo + 32
                  Title = "Atenção !!!"
                  Help = "DEMO.HLP"
                  Ctxt = 1000
                  RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                  If RESPOSTA = vbYes Then
                        TabCABECA!STATUS = 9
                        TabTemp!dt_cancelamento = Date
                        TabTemp!STATUS = "C"
                     TabTemp.Update
                     'VOLTANDO PRODUTO PARA ESTOQUE
                     If TabCABECA!TIPO_REGISTRO = "R" Then
                        Msg = "Deseja baixar retido ?"
                        Style = vbYesNo + 32
                        Title = "Atenção !!!"
                        Help = "DEMO.HLP"
                        Ctxt = 1000
                        RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                        If RESPOSTA = vbYes Then
                           SQL = "select * from PEDIDOITEM "
                           SQL = SQL & " where numr_req = " & txtOR.Text
                           Set TABREQITEM = DBARQEMP.OpenRecordset(SQL)
                           While Not TABREQITEM.EOF
                              SQL = "update PRODUTO "
                              SQL = SQL & "set qtd_balcao = qtd_balcao - " & TABREQITEM!QTD_PEDIDA
                              SQL = SQL & " where codg_prod = '" & TABREQITEM!CODG_PROD & "'"
                              DBARQEMP.Execute SQL
                              TABREQITEM.MoveNext
                           Wend
                           TABREQITEM.Close
                        End If
                        'LANÇAMENTOS
                        SQL = "select * from LANCAMENTO "
                        SQL = SQL & " where numr_doc = " & txtOR.Text
                        Set TabLancamento = DBARQEMP.OpenRecordset(SQL)
                        If Not TabLancamento.EOF Then
                           Msg = "Deseja Cancelar Lançamento Contas a Receber?"
                           Style = vbYesNo + 32
                           Title = "Atenção !!!"
                           Help = "DEMO.HLP"
                           Ctxt = 1000
                           RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)
                           If RESPOSTA = vbYes Then
                              SQL = "delete * from CAIXADIAITEM "
                              SQL = SQL & " where numr_doc = " & txtOR.Text
                              SQL = SQL & " and caixa_id = " & CAIXA_ID
                              DBARQEMP.Execute SQL
                              
                                 TabLancamento!DT_BAIXA = Date
                                 TabLancamento!TIPO_LANCAMENTO = 9
                              TabLancamento.Update
                              SQL = "select * from ITEMLANCAMENTO "
                              SQL = SQL & " where numr_doc = " & txtOR.Text
                              Set TabTemp = DBARQEMP.OpenRecordset(SQL)
                              While Not TabTemp.EOF

                                    TabTemp!usu_alt = Codg_Usu_N
                                    TabTemp!dt_alt = Date
                                    If IsNull(TabTemp!usu_cad) Then
                                       TabTemp!usu_cad = Codg_Usu_N
                                       TabTemp!DT_CAD = Date
                                       Else
                                          If TabTemp!usu_cad = "" Then
                                             TabTemp!usu_cad = Codg_Usu_N
                                             TabTemp!DT_CAD = Date
                                             Else
                                                If TabTemp!usu_cad <= 0 Then
                                                   TabTemp!usu_cad = Codg_Usu_N
                                                   TabTemp!DT_CAD = Date
                                                End If
                                          End If
                                    End If
                                    TabTemp!DT_BAIXA = Date
                                    TabTemp!STATUS = "C"
                                    TabTemp!CODG_USU_BAIXA = Codg_Usu_N
                                 TabTemp.Update
                                 TabTemp.MoveNext
                              Wend
                              TabTemp.Close
                           End If
                           TabLancamento.Close
                        End If
                     End If
                     If TabCABECA!TIPO_REGISTRO = "S" Then RESPOSTA = "Ordem de Serviço"
                     MsgBox RESPOSTA & " foi cancelado com sucesso."
                  End If
            End If
            Else
               TabCABECA.Close
               MsgBox "Ordem de Serviço inexistente !!!"
               txtOR.SetFocus
         End If
         TabCABECA.Close
         LIMPA_CANC
      Case "limpar"
         LIMPA_CANC
      Case "voltar"
         Unload Me
   End Select
End Sub

Private Sub txtOR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0

      If txtOR.Text = "" Then _
         Exit Sub

      If IsNull(txtOR.Text) Then _
         Exit Sub

      PEDIDO_ID_N = txtOR.Text
      
      SQL = "select * from CABECAOS "
      SQL = SQL & " where pedido_id =  " & PEDIDO_ID_N
      Set TabTemp = DBARQAUX.OpenRecordset(SQL, 4)
      If Not TabTemp.EOF Then
         txtPlaca.Text = Left(TabTemp!PLACA, 3) & "-" & Right(TabTemp!PLACA, 5)

         If Not IsNull(TabTemp!DT_ABERTURA) Then
            DTEMIS.PromptInclude = False
               DTEMIS.Text = TabTemp!DT_ABERTURA
            DTEMIS.PromptInclude = True
         End If

         If Not IsNull(TabTemp!dt_cancelamento) Then
            DTCANCELA.PromptInclude = False
               DTCANCELA.Text = TabTemp!dt_cancelamento
            DTCANCELA.PromptInclude = True
         End If

         SQL = "select * from OSVEICULO "
         SQL = SQL & " where placa = '" & Trim(TabTemp!PLACA) & "'"
         Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
         If Not TabAUX.EOF Then
            'txtPLACA.Text = TABAUX!placa
            CGCCPF.PromptInclude = False
            CGCCPF.Text = TabAUX!CGCCPF

            SQL = "select nome from CLIENTE "
            SQL = SQL & " where cgccpf = '" & TabAUX!CGCCPF & "'"
            Set TabCli = DBARQEMP.OpenRecordset(SQL, 4)
            If Not TabCli.EOF Then
               txtNome.Text = TabCli!NOME
               Else: MsgBox "Cliente não cadastrado, verifique."
            End If
         End If
         TabAUX.Close

          'peças
          NUMR_SEQ_N = 0
          VALOR_DESCONTO_N = 0
          TOTAL_PEÇAS_N = 0
          TOTAL_DESCONTO_PEÇAS_N = 0
    
          SQL = "select * from PEDIDOITEM "
          SQL = SQL & " where numr_req = " & txtOR.Text
          Set TabCABECA = DBARQEMP.OpenRecordset(SQL)
          LISTAPEÇA.ListItems.Clear
          While Not TabCABECA.EOF
             TOTAL_PEÇAS_N = TOTAL_PEÇAS_N + (TabCABECA!Valor_Item * TabCABECA!QTD_PEDIDA)
             TOTAL_DESCONTO_PEÇAS_N = TOTAL_DESCONTO_PEÇAS_N + TabCABECA!Valor_Desconto
            
             Set item = LISTAPEÇA.ListItems.Add(, "seq." & TabCABECA!CODG_PROD, TabCABECA!CODG_PROD)

             SQL = "select descricao,referencia from PRODUTO "
             SQL = SQL & " where codg_prod = '" & TabCABECA!CODG_PROD & "'"
             Set TabUSU = DBARQEMP.OpenRecordset(SQL, dbOpenSnapshot)
             If Not TabUSU.EOF Then
                item.SubItems(1) = TabUSU!DESCRICAO
                If Not IsNull(TabUSU!REFERENCIA) Then _
                   item.SubItems(6) = TabUSU!REFERENCIA
             End If
             TabUSU.Close
             item.SubItems(2) = TabCABECA!QTD_PEDIDA
             item.SubItems(3) = Format(TabCABECA!Valor_Item, strFormatacao2Digitos)
             item.SubItems(4) = Format(TabCABECA!Valor_Desconto, strFormatacao2Digitos)
             item.SubItems(5) = Format(TabCABECA!Valor_Item * TabCABECA!QTD_PEDIDA - TabCABECA!Valor_Desconto, strFormatacao2Digitos)
    
             'txtTOTALPRODUTO.Text = FORMAT(TOTAL_PEÇAS_N - TOTAL_DESCONTO_PEÇAS_N,strFormatacao2Digitos)
             'txtTOTALPRODUTO.Refresh
             'txtDESCONTOPRODUTO.Text = FORMAT(TOTAL_DESCONTO_PEÇAS_N,strFormatacao2Digitos)
             'txtDESCONTOPRODUTO.Refresh
    
             TabCABECA.MoveNext
          Wend
          TabCABECA.Close
          
          'serviços
          LISTASERVIÇO.ListItems.Clear
          VALOR_TOTAL_N = 0
          TOTAL_SERVIÇO_N = 0
          VALOR_DESCONTO_N = 0
          TOTAL_DESCONTO_SERVIÇO_N = 0
          NUMR_SEQ_N = 1
    
          SQL = "select * from ITEMOS "
          SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
          SQL = SQL & " order by hora_inicio"
          Set TABREQITEM = DBARQAUX.OpenRecordset(SQL, 4)
          While Not TABREQITEM.EOF
             NUMR_SEQ_N = 1 + NUMR_SEQ_N
             Set item = LISTASERVIÇO.ListItems.Add(, "seq." & NUMR_SEQ_N, TABREQITEM!OSTAREFA_ID)

             SQL = "select * from OSTAREFA "
             SQL = SQL & " where OSTAREFA_ID = '" & TABREQITEM!OSTAREFA_ID & "'"
             Set TabAUX = DBARQAUX.OpenRecordset(SQL, 4)
             If Not TabAUX.EOF Then _
                item.SubItems(1) = TabAUX!DESCRICAO
             TabAUX.Close

             TOTAL_SERVIÇO_N = TOTAL_SERVIÇO_N + TABREQITEM!valor_tarefa
             TOTAL_DESCONTO_SERVIÇO_N = TOTAL_DESCONTO_SERVIÇO_N + TABREQITEM!valor_desc_tarefa
    
             item.SubItems(2) = Format(TABREQITEM!valor_tarefa, strFormatacao2Digitos)
             item.SubItems(3) = Format(TABREQITEM!valor_desc_tarefa, strFormatacao2Digitos)
             item.SubItems(4) = Format(TABREQITEM!valor_tarefa - TABREQITEM!valor_desc_tarefa, strFormatacao2Digitos)
             If TABREQITEM!STATUS = "A" Then _
                item.SubItems(5) = "Ativo"
             If TABREQITEM!STATUS = "B" Then _
                item.SubItems(5) = "Baixado"
             If TABREQITEM!STATUS = "C" Then _
                item.SubItems(5) = "Cancelado"
             If TABREQITEM!STATUS = "E" Then _
                item.SubItems(5) = "Execução"
    
             SQL = "select * from USUARIO "
             SQL = SQL & " where codigo = " & TABREQITEM!codg_mecanico
             Set TabDESCR = DBARQEMP.OpenRecordset(SQL, 4)
             If Not TabDESCR.EOF Then _
                item.SubItems(6) = TabDESCR!NOME & " - " & TabDESCR!Codigo
             TabDESCR.Close
             TABREQITEM.MoveNext
          Wend
          TABREQITEM.Close
          'txtTOTALSERVIÇO.Text = FORMAT(TOTAL_SERVIÇO_N - TOTAL_DESCONTO_SERVIÇO_N,strFormatacao2Digitos)
          'txtTOTALSERVIÇO.Refresh
          'txtDESCONTOSERVIÇO.Text = FORMAT(TOTAL_DESCONTO_SERVIÇO_N,strFormatacao2Digitos)
          'txtDESCONTOSERVIÇO.Refresh
         If Not IsNull(TabTemp!STATUS) Then
            If TabTemp!STATUS = "A" Then _
               txtStatus.Text = "Aberta"
            If TabTemp!STATUS = "B" Then _
               txtStatus.Text = "BAIXADA"
            If TabTemp!STATUS = "C" Then _
               txtStatus.Text = "CANCELADA"
            If TabTemp!STATUS = "D" Then _
               txtStatus.Text = "NEGOCIAÇÂO"
            If TabTemp!STATUS = "E" Then _
               txtStatus.Text = "EXECUSÃO"
            If TabTemp!STATUS = "F" Then
               txtStatus.Text = "FECHADA"
               MsgBox "Não é permitido cancelar Ordem de Serviço fechada."
               Unload Me
            End If
         End If
      End If
   End If
End Sub

Private Sub LIMPA_CANC()
   txtPlaca.Text = ""
   LISTASERVIÇO.ListItems.Clear
   LISTAPEÇA.ListItems.Clear
   txtOR.Text = ""
   DTEMIS.Text = ""
   DTCANCELA.Text = ""
   txtForma.Text = ""
   CGCCPF.Text = ""
   txtNome.Text = ""
   txtStatus.Text = ""
   txtOR.SetFocus
End Sub
