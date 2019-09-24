VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmINTEGRA 
   Caption         =   "Integração NFe"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INTEGRA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   9495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "ITENS MFI010"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "INTEGRA FINANCEIRO"
      Height          =   405
      Left            =   6480
      TabIndex        =   6
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdTRANSP 
      Caption         =   "Transportadora"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdPRODUTO 
      Caption         =   "Produto"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdIBGE 
      Caption         =   "Atualiza MTACIDADE GLOBAL"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INTEGRA PEDIDO"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IBGE antigo"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCLIENTE 
      Caption         =   "Cliente"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin MSComctlLib.ListView listaMEGASIM 
      Height          =   2385
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   4207
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
      ForeColor       =   4194304
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IBGE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CIDADE"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UF"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CEP"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView listaGLOBAL 
      Height          =   2385
      Left            =   0
      TabIndex        =   8
      Top             =   4800
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   4207
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
      ForeColor       =   4194304
      BackColor       =   16777215
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IBGE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CIDADE"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UF"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CEP"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmINTEGRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim CNPJ_CPF_A    As String
   Dim Conta_N       As Long
'===============CLIENTE
   Dim A1_NOME       As String
   Dim A1_PESSOA     As String
   Dim A1_NREDUZ     As String
   Dim A1_TIPO       As String
   Dim A1_END        As String
   Dim A1_MUN        As String
   Dim A1_EST        As String
   Dim A1_BAIRRO     As String
   Dim A1_ESTADO     As String
   Dim A1_CEP        As String
   Dim A1_DDI        As String
   Dim A1_DDD        As String
   Dim A1_TEL        As String
   Dim A1_TELEX      As String
   Dim A1_FAX        As String
   Dim A1_CGC        As String
   Dim A1_CONTATO    As String
   Dim A1_INSCR      As String
   Dim A1_INSCRM     As String
   Dim A1_PFISICA    As String
   Dim A1_RG         As String
   Dim A1_EMAIL      As String
   Dim A1_HPAGE      As String
   Dim A1_INSCRUR    As String
   Dim A1_CODLOJA    As String
   Dim A1_CODCIDENT  As Long
   Dim A1_FILIAL     As String
   Dim A1_LOJA       As String
'===============PRODUTO
   Dim B1_COD        As String
   Dim B1_DESC       As String
   Dim B1_CODITE     As String
   Dim B1_TIPO       As String
   Dim B1_UM         As String
   Dim B1_LOCPAD     As String
   Dim B1_GRUPO      As String
   Dim B1_POSIPI     As String
   Dim B1_ESPECIE    As String
   Dim B1_EX_NCM     As String
   Dim B1_EX_NBM     As String
   Dim B1_PESO       As String
   Dim B1_PICM       As String
   Dim B1_IPI        As String
   Dim B1_ALIQISS    As String
   Dim B1_CODISS     As String
   Dim B1_TE         As String
   Dim B1_TS         As String
   Dim B1_BITMAP     As String
   Dim B1_SEGUM      As String
   Dim B1_PICMRET    As String
   Dim B1_CONV       As String
   Dim B1_PICMENT    As String
   Dim B1_IMPZFRC    As String
   Dim B1_TIPCONV    As String
   Dim B1_ALTER      As String
   Dim B1_QE         As String
   Dim B1_PRV1       As String
   Dim B1_EMIN       As String
   Dim B1_UCOM       As String
   Dim B1_CUSTD      As String
   Dim B1_MCUSTD     As String
   Dim B1_ESTFOR     As String
   Dim B1_UPRC       As String
   Dim B1_ESTSEG     As String
   Dim B1_FORPRZ     As String
   Dim B1_PE         As String
   Dim B1_TIPE       As String
   Dim B1_LE         As String
   Dim B1_LM         As String
   Dim B1_CONTA      As String
   Dim B1_CC         As String
   Dim B1_TOLER      As String
   Dim B1_ITEMCC     As String
   Dim B1_LOJPROC    As String
   Dim B1_FAMILIA    As String
   Dim B1_QB         As String
   Dim B1_PROC       As String
   Dim B1_APROPRI    As String
   Dim B1_FANTASM    As String
   Dim B1_TIPODEC    As String
   Dim B1_ORIGEM     As String
   Dim B1_DATREF     As String
   Dim B1_CLASFIS    As String
   Dim B1_UREV       As String
   Dim B1_RASTRO     As String
   Dim B1_FORAEST    As String
   Dim B1_COMIS      As String
   Dim B1_MONO       As String
   Dim B1_MRP        As String
   Dim B1_DTREFP1    As String
   Dim B1_PERINV     As String
   Dim B1_GRTRIB     As String
   Dim B1_NOTAMIN    As String
   Dim B1_PRVALID    As String
   Dim B1_CONTSOC    As String
   Dim B1_CONINI     As String
   Dim B1_NUMCOP     As String
   Dim B1_CODBAR     As String
   Dim B1_GRADE      As String
   Dim B1_FORMLOT    As String
   Dim B1_IRRF       As String
   Dim B1_FPCOD      As String
   Dim B1_LOCALIZ    As String
   Dim B1_CONTRAT    As String
   Dim B1_DESC_P     As String
   Dim B1_DESC_GI    As String
   Dim B1_DESC_I     As String
   Dim B1_OPERPAD    As String
   Dim B1_IMPORT     As String
   Dim B1_OPC        As String
   Dim B1_ANUENTE    As String
   Dim B1_CODOBS     As String
   Dim B1_VLREFUS    As String
   Dim B1_FABRIC     As String
   Dim B1_SITPROD    As String
   Dim B1_MODELO     As String
   Dim B1_SETOR      As String
   Dim B1_BALANCA    As String
   Dim B1_PRODPAI    As String
   Dim B1_TECLA      As String
   Dim B1_TIPOCQ     As String
   Dim B1_SOLICIT    As String
   Dim B1_GRUPCOM    As String
   Dim B1_NUMCQPR    As String
   Dim B1_CONTCQP    As String
   Dim B1_REVATU     As String
   Dim B1_INSS       As String
   Dim B1_CODEMB     As String
   Dim B1_ESPECIF    As String
   Dim B1_MAT_PRI    As String
   Dim B1_NALNCCA    As String
   Dim B1_REDINSS    As String
   Dim B1_NALSH      As String
   Dim B1_ALADI      As String
   Dim B1_REDIRRF    As String
   Dim B1_TAB_IPI    As String
   Dim B1_DATASUB    As String
   Dim B1_GRUDES     As String
   Dim B1_PCSLL      As String
   Dim B1_PCOFINS    As String
   Dim B1_PPIS       As String
   Dim B1_MTBF       As String
   Dim B1_REDPIS     As String
   Dim B1_REDCOF     As String
   Dim B1_MTTR       As String
   Dim B1_FLAGSUG    As String
   Dim B1_CLASSVE    As String
   Dim B1_MIDIA      As String
   Dim B1_QTMIDIA    As String
   Dim B1_VLR_IPI    As String
   Dim B1_QTDSER     As String
   Dim B1_ENVOBR     As String
   Dim B1_SERIE      As String
   Dim B1_FAIXAS     As String
   Dim B1_NROPAG     As String
   Dim B1_ISBN       As String
   Dim B1_TITORIG    As String
   Dim B1_LINGUA     As String
   Dim B1_EDICAO     As String
   Dim B1_OBSISBN    As String
   Dim B1_CLVL       As String
   Dim B1_ATIVO      As String
   Dim B1_PESBRU     As String
   Dim B1_TIPCAR     As String
   Dim B1_VLR_ICM    As String
   Dim B1_VLRSELO    As String
   Dim B1_CODNOR     As String
   Dim B1_CORPRI     As String
   Dim B1_CORSEC     As String
   Dim B1_ATRIB1     As String
   Dim B1_ATRIB2     As String
   Dim B1_ATRIB3     As String
   Dim B1_REGSEQ     As String
   Dim B1_NICONE     As String
   Dim B1_UCALSTD    As String
   Dim B1_CPOTENC    As String
   Dim B1_POTENCI    As String
   Dim B1_QTDACUM    As String
   Dim B1_QTDINIC    As String
   Dim B1_REQUIS     As String
   Dim B1_CODMOD     As String
   Dim B1_DESENHO    As String
   Dim B1_CUSTIND    As String
   Dim B1_CUSREAL    As String
   Dim B1_QTDPM      As String
   Dim B1_QTDPP      As String
   Dim B1_PIS        As String
   Dim B1_COFINS     As String
   Dim B1_CSLL       As String
   Dim D_E_L_E_T_    As String
   Dim B1_QTDMIN     As String
   Dim B1_ENDFIS     As String
   Dim B1_NOMEFOTO   As String
   Dim B1_REFERENCIA As String
   Dim B1_CODNCM     As String
'===============CABEÇA NOTA MFA010
   Dim Numr_Nota_N   As String
   Dim MFADOC        As String
   Dim MFASERIE      As String
   Dim MFACLIENTE    As String
   Dim MFACOND       As String
   Dim MFADUPL       As String
   Dim MFAEMISSAO    As String
   Dim MFAEST        As String
   Dim MFATIPOCLI    As String

   Dim MFANFORI      As String
   Dim MFADESCONT    As String
   Dim MFASERIORI    As String
   Dim MFAESPECI1    As String
   Dim MFAESPECI2    As String
   Dim MFAESPECI3    As String
   Dim MFAESPECI4    As String
   Dim MFAVOLUME2    As String
   Dim MFAVOLUME3    As String
   Dim MFAVOLUME4    As String
   Dim MFAICMSRET    As String
   Dim MFAREDESP     As String
   Dim MFAVEND1      As String
   Dim MFAVEND2      As String
   Dim MFAVEND3      As String
   Dim MFAVEND4      As String
   Dim MFAVEND5      As String
   Dim MFAOK         As String
   Dim MFAFIMP       As String
   Dim MFAFATORB0    As String
   Dim MFAFATORB1    As String
   Dim MFAVARIAC     As String
   Dim MFABASEISS    As String
   Dim MFAVALISS     As String
   Dim MFAVALFAT     As String
   Dim MFACONTSOC    As String
   Dim MFABRICMS     As String
   Dim MFAFRETAUT    As String
   Dim MFAICMAUTO    As String
   Dim MFADESPESA    As String
   Dim MFANEXTDOC    As String
   Dim MFAPDV        As String
   Dim MFAMAPA       As String
   Dim MFAECF        As String
   Dim MFABASIMP1    As String
   Dim MFABASIMP2    As String
   Dim MFABASIMP3    As String
   Dim MFABASIMP4    As String
   Dim MFABASIMP5    As String
   Dim MFABASIMP6    As String
   Dim MFAVALIMP1    As String
   Dim MFAVALIMP2    As String
   Dim MFAVALIMP3    As String
   Dim MFAVALIMP4    As String
   Dim MFAVALINSS    As String
   Dim MFAHORA       As String
   Dim MFAMOEDA      As String
   Dim MFAREGIAO     As String
   Dim MFAVALCSLL    As String
   Dim MFAVALCOFI    As String
   Dim MFAVALPIS     As String
   Dim MFALOTE       As String
   Dim MFATXMOEDA    As String
   Dim MFAVALIRRF    As String
   Dim MFACARGA      As String
   Dim MFASEQCAR     As String
   Dim MFANEXTSER    As String
   Dim MFAPEDPEND    As String
   Dim MFADESCCAB    As String
   Dim MFAFORMUL     As String
   Dim MFATIPODOC    As String
   Dim MFANFEACRS    As String
   Dim MFASEQENT     As String
   Dim MFADELETE     As String
   Dim MFAREGISTRO   As String
   Dim MFANFESERVI   As String
   Dim MFANFEHRSE    As String
   Dim MFANFECVS     As String
   Dim MFACODPROT    As String
   Dim MFACODSTAT    As String
   Dim MFACODMORE    As String
   Dim MFACHAVENFE   As String
   Dim MFAMOTRESU    As String
   Dim MFACODRECI    As String
   Dim MFALOTENFE    As String
   Dim MFAPLACA      As String
   Dim MFAUFPLACA    As String
   Dim MFAINDPAG     As String
   Dim MFAVALTOT     As String
   Dim MFAVALLIQUI   As String
'===============ITENS NOTA PRODUTO MFI010
      Dim MFICOD        As String
      Dim MFIUM         As String
      Dim MFISEGUM      As String
      Dim MFIQUANT      As String
      Dim MFIPRCVEN     As String
      Dim MFITOTAL      As String
      Dim MFIVALIPI     As String
      Dim MFIVALICM     As String
      Dim MFITES        As String
      Dim MFIDESC       As String
      Dim MFIIPI        As String
      Dim MFIPICM       As String
      Dim MFIPESO       As String
      Dim MFICONTA      As String
      Dim MFIOP         As String
      Dim MFIITEMPV     As String
      Dim MFIGRUPO      As String
      Dim MFITP         As String
      Dim MFICUSTO1     As String
      Dim MFICUSTO2     As String
      Dim MFICUSTO3     As String
      Dim MFICUSTO4     As String
      Dim MFICUSTO5     As String
      Dim MFIPRUNIT     As String
      Dim MFIQTSEGUM    As String
      Dim MFIEST        As String
      Dim MFIDESCON     As String
      Dim MFITIPO       As String
      Dim MFINFORI      As String
      Dim MFISERIORI    As String
      Dim MFIQTDEDEV    As String
      Dim MFIVALDEV     As String
      Dim MFIORIGLAN    As String
      Dim MFIBRICMS     As String
      Dim MFIBASEORI    As String
      Dim MFIBASEICM    As String
      Dim MFIVALACRS    As String
      Dim MFIIDENTB6    As String
      Dim MFICODISS     As String
      Dim MFIGRADE      As String
      Dim MFISEQCALC    As String
      Dim MFIICMSRET    As String
      Dim MFICOMIS1     As String
      Dim MFICOMIS2     As String
      Dim MFICOMIS3     As String
      Dim MFICOMIS4     As String
      Dim MFICOMIS5     As String
      Dim MFILOTECTL    As String
      Dim MFINUMLOTE    As String
      Dim MFIDTVALID    As String
      Dim MFIDESCZFR    As String
      Dim MFIPDV        As String
      Dim MFINUMSERI    As String
      Dim MFIDTLCTCT    As String
      Dim MFICUSFF1     As String
      Dim MFICUSFF2     As String
      Dim MFICUSFF3     As String
      Dim MFICUSFF4     As String
      Dim MFICUSFF5     As String
      Dim MFIBASIMP1    As String
      Dim MFIBASIMP2    As String
      Dim MFIBASIMP3    As String
      Dim MFIBASIMP4    As String
      Dim MFIBASIMP5    As String
      Dim MFIBASIMP6    As String
      Dim MFIVALIMP1    As String
      Dim MFIVALIMP2    As String
      Dim MFIVALIMP3    As String
      Dim MFIVALIMP4    As String
      Dim MFIVALIMP5    As String
      Dim MFIVALIMP6    As String
      Dim MFIITEMORI    As String
      Dim MFICODFAB     As String
      Dim MFILOJAFA     As String
      Dim MFICCUSTO     As String
      Dim MFIITEMCC     As String
      Dim MFILOCALIZ    As String
      Dim MFIENVCNAB    As String
      Dim MFIALIQINS    As String
      Dim MFIPREEMB     As String
      Dim MFIALIQISS    As String
      Dim MFIBASEIPI    As String
      Dim MFIBASEISS    As String
      Dim MFIVALISS     As String
      Dim MFISEGURO     As String
      Dim MFIVALFRE     As String
      Dim MFIDESPESA    As String
      Dim MFICLVL       As String
      Dim MFIBASEINS    As String
      Dim MFIICMFRET    As String
      Dim MFISERVIC     As String
      Dim MFISTSERV     As String
      Dim MFIVALINS     As String
      Dim MFIPROJPMS    As String
      Dim MFITASKPMS    As String
      Dim MFILICITA     As String
      Dim MFIREMITO     As String
      Dim MFISERIREM    As String
      Dim MFIITEMREM    As String
      Dim MFIALQIMP1    As String
      Dim MFIALQIMP2    As String
      Dim MFIALQIMP3    As String
      Dim MFIALQIMP4    As String
      Dim MFIALQIMP5    As String
      Dim MFIALQIMP6    As String
      Dim MFITPDCENV    As String
      Dim MFIOK         As String
      Dim MFIENDER      As String
      Dim MFIEDTPMS     As String
      Dim MFIVARPRUN    As String
      Dim MFIFORMUL     As String
      Dim MFITIPODOC    As String
      Dim MFIVAC        As String
      Dim MFITIPOREM    As String
      Dim MFIQTDEFAT    As String
      Dim MFIQTDAFAT    As String
      Dim MFIPOTENCI    As String
      Dim MFIDELETE     As String
      Dim MFIDESTOTIT   As String
      Dim MFIALIICMS    As String
      Dim MFIBASICMST   As String
      Dim MFIALIICMST   As String
      Dim MFIVALICMST   As String
      Dim MFIALIICMRED  As String
      Dim MFIVALBRUT    As String
      Dim MFIVALBONI    As String
      Dim MFIVALTROCA   As String
      Dim MFIQTDVOL     As String
      Dim MFIPESLIQ     As String
      Dim MFIPESBRU     As String
      Dim MFIVALLIQ     As String
   
Private Sub cmdCLIENTE_Click()
   CLIENTE_INTEGRA ("")
   MsgBox "ok"
End Sub

Private Sub cmdIBGE_Click()
'On Error GoTo ERRO_TRATA

   INTEGRA_IBGE ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "cmdIBGE_Click"
End Sub

Private Sub cmdTRANSP_Click()
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select TRANSP_ID, TRANSPORTADORA.PESSOA_ID, TRANSPORTADORA.STATUS, CONTATO, "
   SQL = SQL & " ESTABELECIMENTO_ID , CNPJCPF, DESCRICAO, RAZAO"
   SQL = SQL & " from TRANSPORTADORA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON TRANSPORTADORA.PESSOA_ID = PESSOA.PESSOA_ID"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Call frmINTEGRA.TRANSPORTADORA_INTEGRA(TabTemp.Fields("cnpjcpf").Value)
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close

MsgBox "OK"
End Sub

Private Sub Command1_Click()
   CONT_N = 0
   NUMR_SEQ_N = 0

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select MTACODCID,MTAUF,MTADESC,MTADATA,MTACODIBGE"
   SQL = SQL & " from MTACIDADE WITH (NOLOCK)"
   SQL = SQL & " where MTACODIBGE > 0"
   SQL = SQL & " order by MTAUF"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      Command1.Caption = Trim(TabTemp.Fields("mtadesc").Value)
      DoEvents

      If TabCEP.State = 1 Then _
         TabCEP.Close

      SQL = "select * from CEP WITH (NOLOCK)"
      SQL = SQL & " where cidade = '" & Trim(TabTemp.Fields("mtadesc").Value) & "'"
      SQL = SQL & " and uf = '" & Trim(TabTemp.Fields("mtauf").Value) & "'"
      TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCEP.EOF Then
         If Trim(TabTemp.Fields("MTACODIBGE").Value) <> Trim(TabCEP.Fields("IBGE_ID").Value) Then

            CONT_N = CONT_N + 1

            Set item = listaMEGASIM.ListItems.Add(, "seq." & CONT_N, TabCEP.Fields("IBGE_ID").Value)
            item.SubItems(1) = TabCEP.Fields("cidade").Value
            item.SubItems(2) = TabCEP.Fields("uf").Value
            item.SubItems(3) = TabCEP.Fields("cep_id").Value

            Set item = listaGLOBAL.ListItems.Add(, "seq." & CONT_N, Trim(TabTemp.Fields("MTACODIBGE").Value))
            item.SubItems(1) = Trim(TabTemp.Fields("mtadesc").Value)
            item.SubItems(2) = Trim(TabTemp.Fields("mtauf").Value)
            item.SubItems(3) = ""

            SQL = "update CEP set IBGE_ID = " & Trim(TabTemp.Fields("MTACODIBGE").Value)
            SQL = SQL & " where cidade = '" & Trim(TabTemp.Fields("mtadesc").Value) & "'"
            SQL = SQL & " and uf = '" & Trim(TabTemp.Fields("mtauf").Value) & "'"
            'CONECTA_RETAGUARDA.Execute SQL
         End If
         Else
            NUMR_SEQ_N = NUMR_SEQ_N + 1
            'SQL = "insert into CEP values("
            '   SQL = SQL & "'" & Trim(TABTEMP.Fields("mtadesc").Value)
            'SQL = SQL & ")"
      End If

      If TabCEP.State = 1 Then _
         TabCEP.Close

      TabTemp.MoveNext
   Wend

   If TabTemp.State = 1 Then _
      TabTemp.Close

   MsgBox "não encontrados = " & NUMR_SEQ_N & "           alterados = " & CONT_N
End Sub

Private Sub Command2_Click()
   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select nf_id,modelo_doc from NF order by nf_id desc"
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabTemp.EOF
      If PEDIDO_INTEGRA_MFA010(TabTemp.Fields(0).Value, _
                               1, _
                               TabTemp.Fields(1).Value, _
                               "000", _
                               "OBSNOTA", _
                               "1", _
                               "1", _
                               "1", _
                               "1", _
                               "1", _
                               "1", _
                               "1", _
                               "1", _
                               "5102", _
                               "1", _
                               "1", _
                               "1", _
                               "N", 0) = True Then

         'MsgBox "Ok passou  MFA010"
         Else: MsgBox "Não passou MFA010"
      End If
      DoEvents
      Command2.Caption = TabTemp.Fields(0).Value
      Command2.Refresh
      TabTemp.MoveNext
   Wend
   If TabTemp.State = 1 Then _
      TabTemp.Close
End Sub

Private Sub cmdProduto_Click()
   INTEGRA_PRODUTO 0
End Sub

Private Sub Command4_Click()

   CONT_N = 0

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabMFA010_FDP     As New ADODB.Recordset
   Dim TabPedidoItem_FDP As New ADODB.Recordset
   Dim MFADOC_A          As String
   Dim MFAPREFIXO_A      As String

   If TabMFA010_FDP.State = 1 Then _
      TabMFA010_FDP.Close

   SQL = "SELECT MFADOC,MFASEQUENCIA,MFAEMISSAO,MFAVALBRUT,MFAVALMERC,MFAPREFIXO,MFAREGISTRO"
   SQL = SQL & " ,MFACODPROT,MFACODSTAT,MFACODMORE,MFACHAVENFE,MFAMOTRESU,MFACODRECI,MFALOTENFE"
   SQL = SQL & " ,MFACODSITT,MFAINDPAG,MFADTENSAI,MFATIFRETE,MFADTDIGIT,MFAVALTOT,MFABASICMST"
   SQL = SQL & " ,MFAVALICMST,MFAVALLIQUI,MFAOBSNOTA,MFANFECNF,MFAEMAILENVIADO,MFAINDFINAL,MFAIDDEST"
   SQL = SQL & " ,MFAINDPRES,MFACHAVEREFNFE,MFAFINNFE,MFANOMECONSUMIDOR,MFACPFCONSUMIDOR,vFCPST,vFCPSTRet,MFACLIENTE"
   SQL = SQL & " From MFA010 WITH (NOLOCK)"

   SQL = SQL & " where MFASEQUENCIA not in (select mfisequen from MFI010)"

SQL = SQL & " order by MFASEQUENCIA "

   TabMFA010_FDP.Open SQL, CONECTA_GLOBAL, , , adCmdText
   While Not TabMFA010_FDP.EOF

      If Not IsNull(TabMFA010_FDP.Fields("mfadoc").Value) Then
         MFADOC_A = "" & TabMFA010_FDP.Fields("mfadoc").Value
         MFAPREFIXO_A = "" & TabMFA010_FDP.Fields("MFAPREFIXO").Value

         If TabPedidoItem_FDP.State = 1 Then _
            TabPedidoItem_FDP.Close

            SQL = "SELECT PEDIDOITEM.PEDIDO_ID, PEDIDOITEM.QTD_PEDIDA * PEDIDOITEM.VALOR_ITEM AS TotalItem, "
            SQL = SQL & " PEDIDOITEM.VALOR_DESCONTO, PRODUTO.CODG_PRODUTO, NFITEM.SEQ_ID, NFITEM.PRODUTO_ID, "
            SQL = SQL & " NFITEM.CFOP_ID, NFITEM.STRIBUTARIA, NFITEM.VLRBASEICMS, NFITEM.PERCICMS, NFITEM.VLRICMS, "
            SQL = SQL & " NFITEM.VLRBASEICMSSUBST, NFITEM.PERCICMSSUBST, NFITEM.VLRICMSSUBST,Nfitem.PERCREDUCAOICMS , "
            SQL = SQL & " Nfitem.PERCIVA, Nfitem.PERC_IPI"
            SQL = SQL & " FROM PEDIDOITEM WITH (NOLOCK)"
            SQL = SQL & " INNER JOIN NF WITH (NOLOCK)"
            SQL = SQL & " ON PEDIDOITEM.PEDIDO_ID = NF.PEDIDO_ID "
            SQL = SQL & " INNER JOIN PRODUTO WITH (NOLOCK)"
            SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID "
            SQL = SQL & " INNER JOIN NFITEM WITH (NOLOCK)"
            SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID "
            SQL = SQL & " AND PRODUTO.PRODUTO_ID = NFITEM.PRODUTO_ID"

SQL = SQL & " where numr_nota = " & MFADOC_A

            TabPedidoItem_FDP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabPedidoItem_FDP.EOF Then

PEDIDO_ID_N = 0 & TabPedidoItem_FDP.Fields("pedido_id").Value

               If PEDIDOitem_INTEGRA_MFI010(TabMFA010_FDP.Fields("MFASEQUENCIA").Value, MFADOC_A, MFAPREFIXO_A, TabMFA010_FDP.Fields("MFACLIENTE").Value, TabMFA010_FDP.Fields("MFAEMISSAO").Value) = True Then
                  FINANCEIRO_INTEGRA MFAPREFIXO_A, MFADOC_A
                  CONT_N = CONT_N + 1
                  Else
                     MsgBox "Não passou MFI010 " & PEDIDO_ID_N
                     TRATA_ERROS "Não gravou item pedido = " & PEDIDO_ID_N, Me.Name, "PEDIDOitem_INTEGRA_MFI010"
               End If
            End If
         If TabPedidoItem_FDP.State = 1 Then _
            TabPedidoItem_FDP.Close

         DoEvents
         Command4.Caption = "" & PEDIDO_ID_N
      End If

      TabMFA010_FDP.MoveNext
   Wend
   If TabMFA010_FDP.State = 1 Then _
      TabMFA010_FDP.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

MsgBox "OK FIM   =  " & CONT_N
End Sub
'=======================
Sub LIMPA_VARIAVEIS()

   A1_NOME = ""
   A1_PESSOA = ""
   A1_NREDUZ = ""
   A1_TIPO = ""
   A1_END = ""
   A1_MUN = ""
   A1_EST = ""
   A1_BAIRRO = ""
   A1_ESTADO = ""
   A1_CEP = ""
   A1_DDI = ""
   A1_DDD = ""
   A1_TEL = ""
   A1_TELEX = ""
   A1_FAX = ""
   A1_CGC = ""
   A1_CONTATO = ""
   A1_INSCR = ""
   A1_INSCRM = ""
   A1_PFISICA = ""
   A1_RG = ""
   A1_EMAIL = ""
   A1_HPAGE = ""
   A1_INSCRUR = ""
   NUMR_ID_N = 0

   CODG_PRODUTO_A = ""

   B1_COD = ""
   B1_DESC = ""
   B1_CODITE = ""
   B1_TIPO = ""
   B1_UM = ""
   B1_LOCPAD = ""
   B1_GRUPO = ""
   B1_POSIPI = ""
   B1_ESPECIE = ""
   B1_EX_NCM = ""
   B1_EX_NBM = ""
   B1_PESO = ""
   B1_PICM = ""
   B1_IPI = ""
   B1_ALIQISS = ""
   B1_CODISS = ""
   B1_TE = ""
   B1_TS = ""
   B1_BITMAP = ""
   B1_SEGUM = ""
   B1_PICMRET = ""
   B1_CONV = ""
   B1_PICMENT = ""
   B1_IMPZFRC = ""
   B1_TIPCONV = ""
   B1_ALTER = ""
   B1_QE = ""
   B1_PRV1 = ""
   B1_EMIN = ""
   B1_UCOM = ""
   B1_CUSTD = ""
   B1_MCUSTD = ""
   B1_ESTFOR = ""
   B1_UPRC = ""
   B1_ESTSEG = ""
   B1_FORPRZ = ""
   B1_PE = ""
   B1_TIPE = ""
   B1_LE = ""
   B1_LM = ""
   B1_CONTA = ""
   B1_CC = ""
   B1_TOLER = ""
   B1_ITEMCC = ""
   B1_LOJPROC = ""
   B1_FAMILIA = ""
   B1_QB = ""
   B1_PROC = ""
   B1_APROPRI = ""
   B1_FANTASM = ""
   B1_TIPODEC = ""
   B1_ORIGEM = ""
   B1_DATREF = ""
   B1_CLASFIS = ""
   B1_UREV = ""
   B1_RASTRO = ""
   B1_FORAEST = ""
   B1_COMIS = ""
   B1_MONO = ""
   B1_MRP = ""
   B1_DTREFP1 = ""
   B1_PERINV = ""
   B1_GRTRIB = ""
   B1_NOTAMIN = ""
   B1_PRVALID = ""
   B1_CONTSOC = ""
   B1_CONINI = ""
   B1_NUMCOP = ""
   B1_CODBAR = ""
   B1_GRADE = ""
   B1_FORMLOT = ""
   B1_IRRF = ""
   B1_FPCOD = ""
   B1_LOCALIZ = ""
   B1_CONTRAT = ""
   B1_DESC_P = ""
   B1_DESC_GI = ""
   B1_DESC_I = ""
   B1_OPERPAD = ""
   B1_IMPORT = ""
   B1_OPC = ""
   B1_ANUENTE = ""
   B1_CODOBS = ""
   B1_VLREFUS = ""
   B1_FABRIC = ""
   B1_SITPROD = ""
   B1_MODELO = ""
   B1_SETOR = ""
   B1_BALANCA = ""
   B1_PRODPAI = ""
   B1_TECLA = ""
   B1_TIPOCQ = ""
   B1_SOLICIT = ""
   B1_GRUPCOM = ""
   B1_NUMCQPR = ""
   B1_CONTCQP = ""
   B1_REVATU = ""
   B1_INSS = ""
   B1_CODEMB = ""
   B1_ESPECIF = ""
   B1_MAT_PRI = ""
   B1_NALNCCA = ""
   B1_REDINSS = ""
   B1_NALSH = ""
   B1_ALADI = ""
   B1_REDIRRF = ""
   B1_TAB_IPI = ""
   B1_DATASUB = ""
   B1_GRUDES = ""
   B1_PCSLL = ""
   B1_PCOFINS = ""
   B1_PPIS = ""
   B1_MTBF = ""
   B1_REDPIS = ""
   B1_REDCOF = ""
   B1_MTTR = ""
   B1_FLAGSUG = ""
   B1_CLASSVE = ""
   B1_MIDIA = ""
   B1_QTMIDIA = ""
   B1_VLR_IPI = ""
   B1_QTDSER = ""
   B1_ENVOBR = ""
   B1_SERIE = ""
   B1_FAIXAS = ""
   B1_NROPAG = ""
   B1_ISBN = ""
   B1_TITORIG = ""
   B1_LINGUA = ""
   B1_EDICAO = ""
   B1_OBSISBN = ""
   B1_CLVL = ""
   B1_ATIVO = ""
   B1_PESBRU = ""
   B1_TIPCAR = ""
   B1_VLR_ICM = ""
   B1_VLRSELO = ""
   B1_CODNOR = ""
   B1_CORPRI = ""
   B1_CORSEC = ""
   B1_ATRIB1 = ""
   B1_ATRIB2 = ""
   B1_ATRIB3 = ""
   B1_REGSEQ = ""
   B1_NICONE = ""
   B1_UCALSTD = ""
   B1_CPOTENC = ""
   B1_POTENCI = ""
   B1_QTDACUM = ""
   B1_QTDINIC = ""
   B1_REQUIS = ""
   B1_CODMOD = ""
   B1_DESENHO = ""
   B1_CUSTIND = ""
   B1_CUSREAL = ""
   B1_QTDPM = ""
   B1_QTDPP = ""
   B1_PIS = ""
   B1_COFINS = ""
   B1_CSLL = ""
   D_E_L_E_T_ = ""
   B1_QTDMIN = ""
   B1_ENDFIS = ""
   B1_NOMEFOTO = ""
   B1_REFERENCIA = ""
   B1_CODNCM = ""

   MFADOC = ""
   MFASERIE = ""
   MFACLIENTE = ""
   MFACOND = ""
   MFADUPL = ""
   MFAEMISSAO = ""
   MFAEST = ""
   MFATIPOCLI = ""
   MFANFORI = ""
   MFADESCONT = ""
   MFASERIORI = ""
   MFAESPECI1 = ""
   MFAESPECI2 = ""
   MFAESPECI3 = ""
   MFAESPECI4 = ""
   MFAVOLUME2 = ""
   MFAVOLUME3 = ""
   MFAVOLUME4 = ""
   MFAICMSRET = ""
   MFAREDESP = ""
   MFAVEND1 = ""
   MFAVEND2 = ""
   MFAVEND3 = ""
   MFAVEND4 = ""
   MFAVEND5 = ""
   MFAOK = ""
   MFAFIMP = ""
   MFAFATORB0 = ""
   MFAFATORB1 = ""
   MFAVARIAC = ""
   MFABASEISS = ""
   MFAVALISS = ""
   MFAVALFAT = ""
   MFACONTSOC = ""
   MFABRICMS = ""
   MFAFRETAUT = ""
   MFAICMAUTO = ""
   MFADESPESA = ""
   MFANEXTDOC = ""
   MFAPDV = ""
   MFAMAPA = ""
   MFAECF = ""
   MFABASIMP1 = ""
   MFABASIMP2 = ""
   MFABASIMP3 = ""
   MFABASIMP4 = ""
   MFABASIMP5 = ""
   MFABASIMP6 = ""
   MFAVALIMP1 = ""
   MFAVALIMP2 = ""
   MFAVALIMP3 = ""
   MFAVALIMP4 = ""
   MFAVALINSS = ""
   MFAHORA = ""
   MFAMOEDA = ""
   MFAREGIAO = ""
   MFAVALCSLL = ""
   MFAVALCOFI = ""
   MFAVALPIS = ""
   MFALOTE = ""
   MFATXMOEDA = ""
   MFAVALIRRF = ""
   MFACARGA = ""
   MFASEQCAR = ""
   MFANEXTSER = ""
   MFAPEDPEND = ""
   MFADESCCAB = ""
   MFAFORMUL = ""
   MFATIPODOC = ""
   MFANFEACRS = ""
   MFASEQENT = ""
   MFADELETE = ""
   MFAREGISTRO = ""
   MFANFESERVI = ""
   MFANFEHRSE = ""
   MFANFECVS = ""
   MFACODPROT = ""
   MFACODSTAT = ""
   MFACODMORE = ""
   MFACHAVENFE = ""
   MFAMOTRESU = ""
   MFACODRECI = ""
   MFALOTENFE = ""
   MFAPLACA = ""
   MFAUFPLACA = ""
   MFAINDPAG = ""
   MFAVALTOT = ""
   MFAVALLIQUI = ""
   MFICOD = ""
   MFIUM = ""
   MFISEGUM = ""
   MFIQUANT = ""
   MFIPRCVEN = ""
   MFITOTAL = ""
   MFIVALIPI = ""
   MFIVALICM = ""
   MFITES = ""
   MFIDESC = ""
   MFIIPI = ""
   MFIPICM = ""
   MFIPESO = ""
   MFICONTA = ""
   MFIOP = ""
   MFIITEMPV = ""
   MFIGRUPO = ""
   MFITP = ""
   MFICUSTO1 = ""
   MFICUSTO2 = ""
   MFICUSTO3 = ""
   MFICUSTO4 = ""
   MFICUSTO5 = ""
   MFIPRUNIT = ""
   MFIQTSEGUM = ""
   MFIEST = ""
   MFIDESCON = ""
   MFITIPO = ""
   MFINFORI = ""
   MFISERIORI = ""
   MFIQTDEDEV = ""
   MFIVALDEV = ""
   MFIORIGLAN = ""
   MFIBRICMS = ""
   MFIBASEORI = ""
   MFIBASEICM = ""
   MFIVALACRS = ""
   MFIIDENTB6 = ""
   MFICODISS = ""
   MFIGRADE = ""
   MFISEQCALC = ""
   MFIICMSRET = ""
   MFICOMIS1 = ""
   MFICOMIS2 = ""
   MFICOMIS3 = ""
   MFICOMIS4 = ""
   MFICOMIS5 = ""
   MFILOTECTL = ""
   MFINUMLOTE = ""
   MFIDTVALID = ""
   MFIDESCZFR = ""
   MFIPDV = ""
   MFINUMSERI = ""
   MFIDTLCTCT = ""
   MFICUSFF1 = ""
   MFICUSFF2 = ""
   MFICUSFF3 = ""
   MFICUSFF4 = ""
   MFICUSFF5 = ""
   MFIBASIMP1 = ""
   MFIBASIMP2 = ""
   MFIBASIMP3 = ""
   MFIBASIMP4 = ""
   MFIBASIMP5 = ""
   MFIBASIMP6 = ""
   MFIVALIMP1 = ""
   MFIVALIMP2 = ""
   MFIVALIMP3 = ""
   MFIVALIMP4 = ""
   MFIVALIMP5 = ""
   MFIVALIMP6 = ""
   MFIITEMORI = ""
   MFICODFAB = ""
   MFILOJAFA = ""
   MFICCUSTO = ""
   MFIITEMCC = ""
   MFILOCALIZ = ""
   MFIENVCNAB = ""
   MFIALIQINS = ""
   MFIPREEMB = ""
   MFIALIQISS = ""
   MFIBASEIPI = ""
   MFIBASEISS = ""
   MFIVALISS = ""
   MFISEGURO = ""
   MFIVALFRE = ""
   MFIDESPESA = ""
   MFICLVL = ""
   MFIBASEINS = ""
   MFIICMFRET = ""
   MFISERVIC = ""
   MFISTSERV = ""
   MFIVALINS = ""
   MFIPROJPMS = ""
   MFITASKPMS = ""
   MFILICITA = ""
   MFIREMITO = ""
   MFISERIREM = ""
   MFIITEMREM = ""
   MFIALQIMP1 = ""
   MFIALQIMP2 = ""
   MFIALQIMP3 = ""
   MFIALQIMP4 = ""
   MFIALQIMP5 = ""
   MFIALQIMP6 = ""
   MFITPDCENV = ""
   MFIOK = ""
   MFIENDER = ""
   MFIEDTPMS = ""
   MFIVARPRUN = ""
   MFIFORMUL = ""
   MFITIPODOC = ""
   MFIVAC = ""
   MFITIPOREM = ""
   MFIQTDEFAT = ""
   MFIQTDAFAT = ""
   MFIPOTENCI = ""
   MFIDELETE = ""
   MFIDESTOTIT = ""
   MFIALIICMS = ""
   MFIBASICMST = ""
   MFIALIICMST = ""
   MFIVALICMST = ""
   MFIALIICMRED = ""
   MFIVALBRUT = ""
   MFIVALBONI = ""
   MFIVALTROCA = ""
   MFIQTDVOL = ""
   MFIPESLIQ = ""
   MFIPESBRU = ""
   MFIVALLIQ = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "LIMPA_VARIAVEIS"
End Sub

Public Function TRAZ_ID_TABELA_GLOBAL(NOME_TABELA As String, NOME_CAMPO As String, Campo1_A As String, Condicao1_A As String) As Long
'On Error GoTo ERRO_TRATA

   'ABRE_BANCO_GLOBAL

   'If CONECTA_GLOBAL.State <> 1 Then
   '   MsgBox "Banco GLOBAL não conectado."
   '   Exit Function
   'End If
   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   TRAZ_ID_TABELA_GLOBAL = 0
   If Trim(NOME_TABELA) <> "" And Trim(NOME_CAMPO) <> "" And Trim(Campo1_A) <> "" And Trim(Condicao1_A) <> "" Then
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select " & NOME_CAMPO & " from  " & NOME_TABELA & " WITH (NOLOCK)"
      SQL = SQL & " where  " & Campo1_A & " = '" & Condicao1_A & "'"
      TabDESCR.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabDESCR.EOF Then _
         If Not IsNull(TabDESCR.Fields(0).Value) Then _
            TRAZ_ID_TABELA_GLOBAL = 0 & Trim(TabDESCR.Fields(0).Value)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If
   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_ID_TABELA_GLOBAL"
End Function

Sub INTEGRA_CFOP(CFOP_N As Integer)
'On Error GoTo ERRO_TRATA

   'ABRE_BANCO_GLOBAL

   'If CONECTA_GLOBAL.State <> 1 Then
   '   MsgBox "Banco GLOBAL não conectado."
   '   Exit Sub
   'End If
   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabMTSITTRIBU As New ADODB.Recordset
   Dim TabCFOP       As New ADODB.Recordset

   If TabMTSITTRIBU.State = 1 Then _
      TabMTSITTRIBU.Close

   SQL = "select * from MTSITTRIBU WITH (NOLOCK)"
   SQL = SQL & " where mtscodfis = " & Trim(CFOP_N)
   TabMTSITTRIBU.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If TabMTSITTRIBU.EOF Then
      If TabCFOP.State = 1 Then _
         TabCFOP.Close

      SQL = "select * from CFOP WITH (NOLOCK)"
      SQL = SQL & " where cfop_id = " & Trim(CFOP_N)
      TabCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCFOP.EOF Then
         NUMR_ID_N = 0

         If TabMTSITTRIBU.State = 1 Then _
            TabMTSITTRIBU.Close

         SQL = "select max(MTSCODIGO) from MTSITTRIBU WITH (NOLOCK)"
         TabMTSITTRIBU.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabMTSITTRIBU.EOF Then _
            If Not IsNull(TabMTSITTRIBU.Fields(0).Value) Then _
               NUMR_ID_N = TabMTSITTRIBU.Fields(0).Value + 1

         SQL = "insert into MTSITTRIBU "
            SQL = SQL & "("
               SQL = SQL & "MTSCODIGO,MTSDESCRI,MTSCODFIS,MTSTIPO,MTSCFOPRE,MTSGERASINT"
            SQL = SQL & ")"
         SQL = SQL & " values ("
            SQL = SQL & NUMR_ID_N                                             'MTSCODIGO
            SQL = SQL & ",'" & Trim(TabCFOP.Fields("descricao").Value) & "'"  'MTSDESCRI
            SQL = SQL & "," & Trim(TabCFOP.Fields("cfop_id").Value)           'MTSCODFIS
            SQL = SQL & ",''"                                                 'MTSTIPO
            SQL = SQL & ",''"                                                 'MTSCFOPRE
            SQL = SQL & ",1"                                                  'MTSGERASINT
         SQL = SQL & ")"

         CONECTA_GLOBAL.Execute SQL
      End If
   End If
   If TabMTSITTRIBU.State = 1 Then _
      TabMTSITTRIBU.Close
   If TabCFOP.State = 1 Then _
      TabCFOP.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INTEGRA_CFOP"
End Sub

Sub INTEGRA_IBGE(CODG_IBGE_A As String)
'On Error GoTo ERRO_TRATA

   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   'ABRE_BANCO_GLOBAL

   'If CONECTA_GLOBAL.State <> 1 Then
      'MsgBox "Banco GLOBAL não conectado."
   '   Exit Sub
   'End If
   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabIBGE       As New ADODB.Recordset
   Dim TabMTACIDADE  As New ADODB.Recordset

   If TabIBGE.State = 1 Then _
      TabIBGE.Close

   SQL = "select * from IBGE WITH (NOLOCK)"
   If Trim(CODG_IBGE_A) <> "" Then _
      SQL = SQL & " where ibge_id = '" & Trim(CODG_IBGE_A) & "'"
   TabIBGE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabIBGE.EOF
      If TabMTACIDADE.State = 1 Then _
         TabMTACIDADE.Close

      SQL = "select * from MTACIDADE WITH (NOLOCK)"
      SQL = SQL & " where MTACODIBGE = '" & Trim(TabIBGE.Fields("ibge_id").Value) & "'"
      TabMTACIDADE.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If TabMTACIDADE.EOF Then
         SQL3 = "" & Replace(TabIBGE.Fields("municipio").Value, "'", " ")
         SQL = "insert into MTACIDADE "
            SQL = SQL & " (MTACODCID,MTAUF,MTADESC,MTADATA,MTACODIBGE)"
         SQL = SQL & " values("
            SQL = SQL & "'" & Trim(TabIBGE.Fields("ibge_id").Value) & "'"    '[MTACODCID]
            SQL = SQL & ",'" & Trim(TabIBGE.Fields("estado").Value) & "'"        '[MTAUF]
            SQL = SQL & ",'" & Trim(SQL3) & "'"                                  '[MTADESC]
            SQL = SQL & ",'" & Now & "'"                                         '[MTADATA]
            SQL = SQL & ",'" & Trim(TabIBGE.Fields("ibge_id").Value) & "'"   '[MTACODIBGE]
         SQL = SQL & " )"
         CONECTA_GLOBAL.Execute SQL
         SQL3 = ""
      End If
      If TabMTACIDADE.State = 1 Then _
         TabMTACIDADE.Close

      TabIBGE.MoveNext
      DoEvents
   Wend
   If TabIBGE.State = 1 Then _
      TabIBGE.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INTEGRA_IBGE"
End Sub

Sub INTEGRA_PRODUTO(PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabProdIntegra   As New ADODB.Recordset
   Dim TabTempIntegra   As New ADODB.Recordset
   Dim INDR_GRAVA_PROD  As Boolean
   Dim B1_CEAN          As String
   Dim B1_CEANTRIB      As String

   If TabProdIntegra.State = 1 Then _
      TabProdIntegra.Close

   SQL = "select * from vwProduto WITH (NOLOCK)"
   SQL = SQL & " WHERE situacao = 'A' "

   If PROD_ID_N > 0 Then _
      SQL = SQL & " and produto_id = " & PROD_ID_N

   SQL = SQL & " ORDER BY descricao"

   TabProdIntegra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProdIntegra.EOF
      INDR_GRAVA_PROD = True
      LIMPA_VARIAVEIS

      If TabProdIntegra.Fields("produto_id").Value > 0 Then
         B1_EX_NCM = "" & Trim(TabProdIntegra.Fields("CODG_NCM").Value)
         If Len(B1_EX_NCM) < 8 Then
            Msg = "Produto : " & Trim(TabProdIntegra.Fields("CODG_PRODUTO").Value) & "-" & Trim(TabProdIntegra.Fields("descricao").Value)
            Msg = Msg & " está com o código NCM incorreto, NCM = " & B1_EX_NCM
            'MsgBox Msg
            INDR_GRAVA_PROD = False
         End If
      End If
      If INDR_GRAVA_PROD = True Then
         '=====================
         A1_FILIAL = "0" & EMPRESA_ID_N
         B1_COD = "" & Trim(TabProdIntegra.Fields("CODG_PRODUTO").Value)
         B1_DESC = "" & Trim(TabProdIntegra.Fields("DESCRICAO").Value)
         B1_CODITE = ""
         B1_TIPO = "MP"
         B1_UM = "" & Trim(Left(TabProdIntegra.Fields("UNIDADE_MEDIDA").Value, 2))
         B1_LOCPAD = "01"
         B1_GRUPO = "" & Trim(Left(TabProdIntegra.Fields("FAMILIAPRODUTO_ID").Value, 4))
         B1_POSIPI = ""
         B1_ESPECIE = "0"
         B1_EX_NCM = "" & Trim(TabProdIntegra.Fields("CODG_NCM").Value)
         B1_EX_NBM = ""
         B1_PESO = "" & tpMOEDA(TabProdIntegra.Fields("PESO_LIQUIDO").Value)
         B1_PICM = "0"
         B1_IPI = "0"
         B1_ALIQISS = "0"
         B1_CODISS = ""
         B1_TE = ""
         B1_TS = ""
         B1_BITMAP = ""
         B1_SEGUM = ""
         B1_PICMRET = "0"
         B1_CONV = "0"
         B1_PICMENT = "0"
         B1_IMPZFRC = ""
         B1_TIPCONV = "M"
         B1_ALTER = ""
         B1_QE = "0"
         B1_PRV1 = "0"
         B1_EMIN = "0"
         B1_UCOM = ""
         B1_CUSTD = ""
         B1_MCUSTD = "1"
         B1_ESTFOR = ""
         B1_UPRC = "" & tpMOEDA(TabProdIntegra.Fields("preco_venda").Value)
         B1_UPRC = "" & Trim(B1_UPRC)
   
         B1_ESTSEG = "0"
         B1_FORPRZ = ""
         B1_PE = "0"
         B1_TIPE = ""
         B1_LE = "0"
         B1_LM = "0"
         B1_CONTA = ""
         B1_CC = ""
         B1_TOLER = "0"
         B1_ITEMCC = ""
         B1_LOJPROC = ""
         B1_FAMILIA = "" '& Trim(TabProdIntegra.Fields("descfamilia").Value)
         B1_QB = "1"
         B1_PROC = ""
         B1_APROPRI = ""
         B1_FANTASM = ""
         B1_TIPODEC = ""
         B1_ORIGEM = "" & Trim(Left(TabProdIntegra.Fields("ORIGEM_MERCADO").Value, 1))
         B1_DATREF = ""
         If Not IsNull(TabProdIntegra.Fields("DT_CADASTRO").Value) Then _
            If Trim(TabProdIntegra.Fields("DT_CADASTRO").Value) <> "" Then _
               B1_DATREF = "" & Year(TabProdIntegra.Fields("DT_CADASTRO").Value) & Month(TabProdIntegra.Fields("DT_CADASTRO").Value) & Day(TabProdIntegra.Fields("DT_CADASTRO").Value)
         B1_CLASFIS = ""
         B1_UREV = ""
         If Not IsNull(TabProdIntegra.Fields("DT_CADASTRO").Value) Then _
            If Trim(TabProdIntegra.Fields("DT_CADASTRO").Value) <> "" Then _
               B1_UREV = "" & Year(TabProdIntegra.Fields("DT_CADASTRO").Value) & Month(TabProdIntegra.Fields("DT_CADASTRO").Value) & Day(TabProdIntegra.Fields("DT_CADASTRO").Value)
         B1_RASTRO = "N"
         B1_FORAEST = ""
         B1_COMIS = "0"
         B1_MONO = ""
         B1_MRP = "S"
         B1_DTREFP1 = ""
         B1_PERINV = "0"
         B1_GRTRIB = ""
         B1_NOTAMIN = "0"
         B1_PRVALID = "0"
         B1_CONTSOC = ""
         B1_CONINI = ""
         B1_NUMCOP = "0"
         B1_CODBAR = "" & TabProdIntegra.Fields("CODG_BARRA").Value
         B1_GRADE = ""
         B1_FORMLOT = ""
         B1_IRRF = ""
         B1_FPCOD = ""
         B1_LOCALIZ = "N"
         B1_CONTRAT = "N"
         B1_DESC_P = ""
         B1_DESC_GI = ""
         B1_DESC_I = ""
         B1_OPERPAD = ""
         B1_IMPORT = "N"
         B1_OPC = ""
         B1_ANUENTE = "2"
         B1_CODOBS = ""
         B1_VLREFUS = "0"
         B1_FABRIC = ""
         B1_SITPROD = "" & TabProdIntegra.Fields("SITUACAO").Value
         B1_MODELO = ""
         B1_SETOR = ""
         B1_BALANCA = ""
         B1_PRODPAI = ""
         B1_TECLA = ""
         B1_TIPOCQ = "M"
         B1_SOLICIT = ""
         B1_GRUPCOM = ""
         B1_NUMCQPR = "0"
         B1_CONTCQP = "0"
         B1_REVATU = ""
         B1_INSS = "N"
         B1_CODEMB = ""
         B1_ESPECIF = ""
         B1_MAT_PRI = ""
         B1_NALNCCA = ""
         B1_REDINSS = "0"
         B1_NALSH = ""
         B1_ALADI = ""
         B1_REDIRRF = "0"
         B1_TAB_IPI = ""
         B1_DATASUB = ""
         B1_GRUDES = ""
         B1_PCSLL = "0"
         B1_PCOFINS = "0"
         B1_PPIS = "0"
         B1_MTBF = "0"
         B1_REDPIS = "0"
         B1_REDCOF = "0"
         B1_MTTR = "0"
         B1_FLAGSUG = ""
         B1_CLASSVE = "1"
         B1_MIDIA = "2"
         B1_QTMIDIA = "0"
         B1_VLR_IPI = "0"
         B1_QTDSER = "1"
         B1_ENVOBR = "0"
         B1_SERIE = ""
         B1_FAIXAS = "0"
         B1_NROPAG = "0"
         B1_ISBN = ""
         B1_TITORIG = ""
         B1_LINGUA = ""
         B1_EDICAO = ""
         B1_OBSISBN = ""
         B1_CLVL = ""
         B1_ATIVO = "S"
         B1_PESBRU = "" & tpMOEDA(TabProdIntegra.Fields("PESO_BRUTO").Value)
         B1_PESBRU = Trim(B1_PESBRU)
         B1_TIPCAR = ""
         B1_VLR_ICM = "0"
         B1_VLRSELO = "0"
         B1_CODNOR = ""
         B1_CORPRI = ""
         B1_CORSEC = ""
         B1_ATRIB1 = ""
         B1_ATRIB2 = ""
         B1_ATRIB3 = ""
         B1_REGSEQ = ""
         B1_NICONE = ""
         B1_UCALSTD = ""
         B1_CPOTENC = ""
         B1_POTENCI = "0"
         B1_QTDACUM = "0"
         B1_QTDINIC = "0"
         B1_REQUIS = ""
         B1_CODMOD = ""
         B1_DESENHO = "-"
         B1_CUSTIND = "0"
         B1_CUSREAL = "0"
         B1_QTDPM = "0"
         B1_QTDPP = "0"
         B1_PIS = ""
         B1_COFINS = ""
         B1_CSLL = ""
         D_E_L_E_T_ = ""
         B1_QTDMIN = "0"
         B1_ENDFIS = ""
         B1_NOMEFOTO = ""
         B1_REFERENCIA = "" & B1_COD   'Trim(TabProdIntegra.Fields("REFERENCIA").Value)
         B1_CODNCM = "" & Trim(TabProdIntegra.Fields("CODG_NCM").Value)
         NUMR_ID_N = 1
         '=====================
         B1_CEAN = "SEM GTIN"
         B1_CEANTRIB = "SEM GTIN"
'====================================

         If TabTempIntegra.State = 1 Then _
            TabTempIntegra.Close
   
         SQL = "select [R_E_C_N_O_] from SB1010 WITH (NOLOCK)"
         SQL = SQL & " where B1_CODANT = '" & Trim(B1_COD) & "'"
         TabTempIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If TabTempIntegra.EOF Then
            NUMR_ID_N = 1
   
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
            SQL = "select max([R_E_C_N_O_]) from SB1010 WITH (NOLOCK)"
            TabConsulta.Open SQL, CONECTA_GLOBAL, , , adCmdText
            If Not TabConsulta.EOF Then _
                If Not IsNull(TabConsulta.Fields(0).Value) Then _
                    NUMR_ID_N = TabConsulta.Fields(0).Value + 1
            If TabConsulta.State = 1 Then _
               TabConsulta.Close
   
            SQL = "insert into SB1010 "
            SQL = SQL & "("
               SQL = SQL & "B1_FILIAL"       '
               SQL = SQL & ",B1_COD"         '[B1_COD]
               SQL = SQL & ",B1_DESC"        '[B1_DESC]
               SQL = SQL & ",B1_CODANT"      '[B1_CODANT]
               SQL = SQL & ",B1_CODITE"
               SQL = SQL & ",B1_TIPO"
               SQL = SQL & ",B1_UM"
               SQL = SQL & ",B1_LOCPAD"
               SQL = SQL & ",B1_GRUPO"
               SQL = SQL & ",B1_POSIPI"
               SQL = SQL & ",B1_ESPECIE"
               SQL = SQL & ",B1_EX_NCM"
               SQL = SQL & ",B1_EX_NBM"
               SQL = SQL & ",B1_PESO"
               SQL = SQL & ",B1_PICM"
               SQL = SQL & ",B1_IPI"
               SQL = SQL & ",B1_ALIQISS"
               SQL = SQL & ",B1_CODISS"
               SQL = SQL & ",B1_TE"
               SQL = SQL & ",B1_TS"
               SQL = SQL & ",B1_BITMAP"
               SQL = SQL & ",B1_SEGUM"
               SQL = SQL & ",B1_PICMRET"
               SQL = SQL & ",B1_CONV"
               SQL = SQL & ",B1_PICMENT"
               SQL = SQL & ",B1_IMPZFRC"
               SQL = SQL & ",B1_TIPCONV"
               SQL = SQL & ",B1_ALTER"
               SQL = SQL & ",B1_QE"
               SQL = SQL & ",B1_PRV1"
               SQL = SQL & ",B1_EMIN"
               SQL = SQL & ",B1_UCOM"
               SQL = SQL & ",B1_CUSTD"
               SQL = SQL & ",B1_MCUSTD"
               SQL = SQL & ",B1_ESTFOR"
               SQL = SQL & ",B1_UPRC"
               SQL = SQL & ",B1_ESTSEG"
               SQL = SQL & ",B1_FORPRZ"
               SQL = SQL & ",B1_PE"
               SQL = SQL & ",B1_TIPE"
               SQL = SQL & ",B1_LE"
               SQL = SQL & ",B1_LM"
               SQL = SQL & ",B1_CONTA"
               SQL = SQL & ",B1_CC"
               SQL = SQL & ",B1_TOLER"
               SQL = SQL & ",B1_ITEMCC"
               SQL = SQL & ",B1_LOJPROC"
               SQL = SQL & ",B1_FAMILIA"
               SQL = SQL & ",B1_QB"
               SQL = SQL & ",B1_PROC"
               SQL = SQL & ",B1_APROPRI"
               SQL = SQL & ",B1_FANTASM"
               SQL = SQL & ",B1_TIPODEC"
               SQL = SQL & ",B1_ORIGEM"
               SQL = SQL & ",B1_DATREF"
               SQL = SQL & ",B1_CLASFIS"
               SQL = SQL & ",B1_UREV"
               SQL = SQL & ",B1_RASTRO"
               SQL = SQL & ",B1_FORAEST"
               SQL = SQL & ",B1_COMIS"
               SQL = SQL & ",B1_MONO"
               SQL = SQL & ",B1_MRP"
               SQL = SQL & ",B1_DTREFP1"
               SQL = SQL & ",B1_PERINV"
               SQL = SQL & ",B1_GRTRIB"
               SQL = SQL & ",B1_NOTAMIN"
               SQL = SQL & ",B1_PRVALID"
               SQL = SQL & ",B1_CONTSOC"
               SQL = SQL & ",B1_CONINI"
               SQL = SQL & ",B1_NUMCOP"
               SQL = SQL & ",B1_CODBAR"
               SQL = SQL & ",B1_GRADE"
               SQL = SQL & ",B1_FORMLOT"
               SQL = SQL & ",B1_IRRF"
               SQL = SQL & ",B1_FPCOD"
               SQL = SQL & ",B1_LOCALIZ"
               SQL = SQL & ",B1_CONTRAT"
               SQL = SQL & ",B1_DESC_P"
               SQL = SQL & ",B1_DESC_GI"
               SQL = SQL & ",B1_DESC_I"
               SQL = SQL & ",B1_OPERPAD"
               SQL = SQL & ",B1_IMPORT"
               SQL = SQL & ",B1_OPC"
               SQL = SQL & ",B1_ANUENTE"
               SQL = SQL & ",B1_CODOBS"
               SQL = SQL & ",B1_VLREFUS"
               SQL = SQL & ",B1_FABRIC"
               SQL = SQL & ",B1_SITPROD"
               SQL = SQL & ",B1_MODELO"
               SQL = SQL & ",B1_SETOR"
               SQL = SQL & ",B1_BALANCA"
               SQL = SQL & ",B1_PRODPAI"
               SQL = SQL & ",B1_TECLA"
               SQL = SQL & ",B1_TIPOCQ"
               SQL = SQL & ",B1_SOLICIT"
               SQL = SQL & ",B1_GRUPCOM"
               SQL = SQL & ",B1_NUMCQPR"
               SQL = SQL & ",B1_CONTCQP"
               SQL = SQL & ",B1_REVATU"
               SQL = SQL & ",B1_INSS"
               SQL = SQL & ",B1_CODEMB"
               SQL = SQL & ",B1_ESPECIF"
               SQL = SQL & ",B1_MAT_PRI"
               SQL = SQL & ",B1_NALNCCA"
               SQL = SQL & ",B1_REDINSS"
               SQL = SQL & ",B1_NALSH"
               SQL = SQL & ",B1_ALADI"
               SQL = SQL & ",B1_REDIRRF"
               SQL = SQL & ",B1_TAB_IPI"
               SQL = SQL & ",B1_DATASUB"
               SQL = SQL & ",B1_GRUDES"
               SQL = SQL & ",B1_PCSLL"
               SQL = SQL & ",B1_PCOFINS"
               SQL = SQL & ",B1_PPIS"
               SQL = SQL & ",B1_MTBF"
               SQL = SQL & ",B1_REDPIS"
               SQL = SQL & ",B1_REDCOF"
               SQL = SQL & ",B1_MTTR"
               SQL = SQL & ",B1_FLAGSUG"
               SQL = SQL & ",B1_CLASSVE"
               SQL = SQL & ",B1_MIDIA"
               SQL = SQL & ",B1_QTMIDIA"
               SQL = SQL & ",B1_VLR_IPI"
               SQL = SQL & ",B1_QTDSER"
               SQL = SQL & ",B1_ENVOBR"
               SQL = SQL & ",B1_SERIE"
               SQL = SQL & ",B1_FAIXAS"
               SQL = SQL & ",B1_NROPAG"
               SQL = SQL & ",B1_ISBN"
               SQL = SQL & ",B1_TITORIG"
               SQL = SQL & ",B1_LINGUA"
               SQL = SQL & ",B1_EDICAO"
               SQL = SQL & ",B1_OBSISBN"
               SQL = SQL & ",B1_CLVL"
               SQL = SQL & ",B1_ATIVO"
               SQL = SQL & ",B1_PESBRU"
               SQL = SQL & ",B1_TIPCAR"
               SQL = SQL & ",B1_VLR_ICM"
               SQL = SQL & ",B1_VLRSELO"
               SQL = SQL & ",B1_CODNOR"
               SQL = SQL & ",B1_CORPRI"
               SQL = SQL & ",B1_CORSEC"
               SQL = SQL & ",B1_ATRIB1"
               SQL = SQL & ",B1_ATRIB2"
               SQL = SQL & ",B1_ATRIB3"
               SQL = SQL & ",B1_REGSEQ"
               SQL = SQL & ",B1_NICONE"
               SQL = SQL & ",B1_UCALSTD"
               SQL = SQL & ",B1_CPOTENC"
               SQL = SQL & ",B1_POTENCI"
               SQL = SQL & ",B1_QTDACUM"
               SQL = SQL & ",B1_QTDINIC"
               SQL = SQL & ",B1_REQUIS"
               SQL = SQL & ",B1_CODMOD"
               SQL = SQL & ",B1_DESENHO"
               SQL = SQL & ",B1_CUSTIND"
               SQL = SQL & ",B1_CUSREAL"
               SQL = SQL & ",B1_QTDPM"
               SQL = SQL & ",B1_QTDPP"
               SQL = SQL & ",B1_PIS"
               SQL = SQL & ",B1_COFINS"
               SQL = SQL & ",B1_CSLL"
               SQL = SQL & ",D_E_L_E_T_"
               SQL = SQL & ",R_E_C_N_O_"
               SQL = SQL & ",B1_QTDMIN"
               SQL = SQL & ",B1_ENDFIS"
               SQL = SQL & ",B1_NOMEFOTO"
               SQL = SQL & ",B1_REFERENCIA"
               SQL = SQL & ",B1_CODNCM"
               SQL = SQL & ",B1_CEAN"
               SQL = SQL & ",B1_CEANTRIB"
            SQL = SQL & ") VALUES("
               SQL = SQL & "'" & A1_FILIAL & "'"      'A1_FILIAL
               SQL = SQL & ",'" & B1_COD & "'"        'B1_COD
               SQL = SQL & ",'" & B1_DESC & "'"        'B1_DESC
               SQL = SQL & ",'" & B1_COD & "'"        'B1_CODANT
               SQL = SQL & ",'" & B1_CODITE & "'"
               SQL = SQL & ",'" & B1_TIPO & "'"
               SQL = SQL & ",'" & B1_UM & "'"
               SQL = SQL & ",'" & B1_LOCPAD & "'"
               SQL = SQL & ",'" & B1_GRUPO & "'"
               SQL = SQL & ",'" & B1_POSIPI & "'"
               SQL = SQL & ",'" & B1_ESPECIE & "'"
               SQL = SQL & ",'" & B1_EX_NCM & "'"
               SQL = SQL & ",'" & B1_EX_NBM & "'"
               SQL = SQL & ",'" & B1_PESO & "'"
               SQL = SQL & ",'" & B1_PICM & "'"
               SQL = SQL & ",'" & B1_IPI & "'"
               SQL = SQL & ",'" & B1_ALIQISS & "'"
               SQL = SQL & ",'" & B1_CODISS & "'"
               SQL = SQL & ",'" & B1_TE & "'"
               SQL = SQL & ",'" & B1_TS & "'"
               SQL = SQL & ",'" & B1_BITMAP & "'"
               SQL = SQL & ",'" & B1_SEGUM & "'"
               SQL = SQL & ",'" & B1_PICMRET & "'"
               SQL = SQL & ",'" & B1_CONV & "'"
               SQL = SQL & ",'" & B1_PICMENT & "'"
               SQL = SQL & ",'" & B1_IMPZFRC & "'"
               SQL = SQL & ",'" & B1_TIPCONV & "'"
               SQL = SQL & ",'" & B1_ALTER & "'"
               SQL = SQL & ",'" & B1_QE & "'"
               SQL = SQL & ",'" & B1_PRV1 & "'"
               SQL = SQL & ",'" & B1_EMIN & "'"
               SQL = SQL & ",'" & B1_UCOM & "'"
               SQL = SQL & ",'" & B1_CUSTD & "'"
               SQL = SQL & ",'" & B1_MCUSTD & "'"
               SQL = SQL & ",'" & B1_ESTFOR & "'"
               SQL = SQL & ",'" & B1_UPRC & "'"
               SQL = SQL & ",'" & B1_ESTSEG & "'"
               SQL = SQL & ",'" & B1_FORPRZ & "'"
               SQL = SQL & ",'" & B1_PE & "'"
               SQL = SQL & ",'" & B1_TIPE & "'"
               SQL = SQL & ",'" & B1_LE & "'"
               SQL = SQL & ",'" & B1_LM & "'"
               SQL = SQL & ",'" & B1_CONTA & "'"
               SQL = SQL & ",'" & B1_CC & "'"
               SQL = SQL & ",'" & B1_TOLER & "'"
               SQL = SQL & ",'" & B1_ITEMCC & "'"
               SQL = SQL & ",'" & B1_LOJPROC & "'"
               SQL = SQL & ",'" & B1_FAMILIA & "'"
               SQL = SQL & ",'" & B1_QB & "'"
               SQL = SQL & ",'" & B1_PROC & "'"
               SQL = SQL & ",'" & B1_APROPRI & "'"
               SQL = SQL & ",'" & B1_FANTASM & "'"
               SQL = SQL & ",'" & B1_TIPODEC & "'"
               SQL = SQL & ",'" & B1_ORIGEM & "'"
               SQL = SQL & ",'" & B1_DATREF & "'"
               SQL = SQL & ",'" & B1_CLASFIS & "'"
               SQL = SQL & ",'" & B1_UREV & "'"
               SQL = SQL & ",'" & B1_RASTRO & "'"
               SQL = SQL & ",'" & B1_FORAEST & "'"
               SQL = SQL & ",'" & B1_COMIS & "'"
               SQL = SQL & ",'" & B1_MONO & "'"
               SQL = SQL & ",'" & B1_MRP & "'"
               SQL = SQL & ",'" & B1_DTREFP1 & "'"
               SQL = SQL & ",'" & B1_PERINV & "'"
               SQL = SQL & ",'" & B1_GRTRIB & "'"
               SQL = SQL & ",'" & B1_NOTAMIN & "'"
               SQL = SQL & ",'" & B1_PRVALID & "'"
               SQL = SQL & ",'" & B1_CONTSOC & "'"
               SQL = SQL & ",'" & B1_CONINI & "'"
               SQL = SQL & ",'" & B1_NUMCOP & "'"
               SQL = SQL & ",'" & B1_CODBAR & "'"
               SQL = SQL & ",'" & B1_GRADE & "'"
               SQL = SQL & ",'" & B1_FORMLOT & "'"
               SQL = SQL & ",'" & B1_IRRF & "'"
               SQL = SQL & ",'" & B1_FPCOD & "'"
               SQL = SQL & ",'" & B1_LOCALIZ & "'"
               SQL = SQL & ",'" & B1_CONTRAT & "'"
               SQL = SQL & ",'" & B1_DESC_P & "'"
               SQL = SQL & ",'" & B1_DESC_GI & "'"
               SQL = SQL & ",'" & B1_DESC_I & "'"
               SQL = SQL & ",'" & B1_OPERPAD & "'"
               SQL = SQL & ",'" & B1_IMPORT & "'"
               SQL = SQL & ",'" & B1_OPC & "'"
               SQL = SQL & ",'" & B1_ANUENTE & "'"
               SQL = SQL & ",'" & B1_CODOBS & "'"
               SQL = SQL & ",'" & B1_VLREFUS & "'"
               SQL = SQL & ",'" & B1_FABRIC & "'"
               SQL = SQL & ",'" & B1_SITPROD & "'"
               SQL = SQL & ",'" & B1_MODELO & "'"
               SQL = SQL & ",'" & B1_SETOR & "'"
               SQL = SQL & ",'" & B1_BALANCA & "'"
               SQL = SQL & ",'" & B1_PRODPAI & "'"
               SQL = SQL & ",'" & B1_TECLA & "'"
               SQL = SQL & ",'" & B1_TIPOCQ & "'"
               SQL = SQL & ",'" & B1_SOLICIT & "'"
               SQL = SQL & ",'" & B1_GRUPCOM & "'"
               SQL = SQL & ",'" & B1_NUMCQPR & "'"
               SQL = SQL & ",'" & B1_CONTCQP & "'"
               SQL = SQL & ",'" & B1_REVATU & "'"
               SQL = SQL & ",'" & B1_INSS & "'"
               SQL = SQL & ",'" & B1_CODEMB & "'"
               SQL = SQL & ",'" & B1_ESPECIF & "'"
               SQL = SQL & ",'" & B1_MAT_PRI & "'"
               SQL = SQL & ",'" & B1_NALNCCA & "'"
               SQL = SQL & ",'" & B1_REDINSS & "'"
               SQL = SQL & ",'" & B1_NALSH & "'"
               SQL = SQL & ",'" & B1_ALADI & "'"
               SQL = SQL & ",'" & B1_REDIRRF & "'"
               SQL = SQL & ",'" & B1_TAB_IPI & "'"
               SQL = SQL & ",'" & B1_DATASUB & "'"
               SQL = SQL & ",'" & B1_GRUDES & "'"
               SQL = SQL & ",'" & B1_PCSLL & "'"
               SQL = SQL & ",'" & B1_PCOFINS & "'"
               SQL = SQL & ",'" & B1_PPIS & "'"
               SQL = SQL & ",'" & B1_MTBF & "'"
               SQL = SQL & ",'" & B1_REDPIS & "'"
               SQL = SQL & ",'" & B1_REDCOF & "'"
               SQL = SQL & ",'" & B1_MTTR & "'"
               SQL = SQL & ",'" & B1_FLAGSUG & "'"
               SQL = SQL & ",'" & B1_CLASSVE & "'"
               SQL = SQL & ",'" & B1_MIDIA & "'"
               SQL = SQL & ",'" & B1_QTMIDIA & "'"
               SQL = SQL & ",'" & B1_VLR_IPI & "'"
               SQL = SQL & ",'" & B1_QTDSER & "'"
               SQL = SQL & ",'" & B1_ENVOBR & "'"
               SQL = SQL & ",'" & B1_SERIE & "'"
               SQL = SQL & ",'" & B1_FAIXAS & "'"
               SQL = SQL & ",'" & B1_NROPAG & "'"
               SQL = SQL & ",'" & B1_ISBN & "'"
               SQL = SQL & ",'" & B1_TITORIG & "'"
               SQL = SQL & ",'" & B1_LINGUA & "'"
               SQL = SQL & ",'" & B1_EDICAO & "'"
               SQL = SQL & ",'" & B1_OBSISBN & "'"
               SQL = SQL & ",'" & B1_CLVL & "'"
               SQL = SQL & ",'" & B1_ATIVO & "'"
               SQL = SQL & ",'" & B1_PESBRU & "'"
               SQL = SQL & ",'" & B1_TIPCAR & "'"
               SQL = SQL & ",'" & B1_VLR_ICM & "'"
               SQL = SQL & ",'" & B1_VLRSELO & "'"
               SQL = SQL & ",'" & B1_CODNOR & "'"
               SQL = SQL & ",'" & B1_CORPRI & "'"
               SQL = SQL & ",'" & B1_CORSEC & "'"
               SQL = SQL & ",'" & B1_ATRIB1 & "'"
               SQL = SQL & ",'" & B1_ATRIB2 & "'"
               SQL = SQL & ",'" & B1_ATRIB3 & "'"
               SQL = SQL & ",'" & B1_REGSEQ & "'"
               SQL = SQL & ",'" & B1_NICONE & "'"
               SQL = SQL & ",'" & B1_UCALSTD & "'"
               SQL = SQL & ",'" & B1_CPOTENC & "'"
               SQL = SQL & ",'" & B1_POTENCI & "'"
               SQL = SQL & ",'" & B1_QTDACUM & "'"
               SQL = SQL & ",'" & B1_QTDINIC & "'"
               SQL = SQL & ",'" & B1_REQUIS & "'"
               SQL = SQL & ",'" & B1_CODMOD & "'"
               SQL = SQL & ",'" & B1_DESENHO & "'"
               SQL = SQL & ",'" & B1_CUSTIND & "'"
               SQL = SQL & ",'" & B1_CUSREAL & "'"
               SQL = SQL & ",'" & B1_QTDPM & "'"
               SQL = SQL & ",'" & B1_QTDPP & "'"
               SQL = SQL & ",'" & B1_PIS & "'"
               SQL = SQL & ",'" & B1_COFINS & "'"
               SQL = SQL & ",'" & B1_CSLL & "'"
               SQL = SQL & ",'" & D_E_L_E_T_ & "'"
               SQL = SQL & ",'" & NUMR_ID_N & "'"
               SQL = SQL & ",'" & B1_QTDMIN & "'"
               SQL = SQL & ",'" & B1_ENDFIS & "'"
               SQL = SQL & ",'" & B1_NOMEFOTO & "'"
               SQL = SQL & ",'" & B1_REFERENCIA & "'"
               SQL = SQL & ",'" & B1_CODNCM & "'"
               SQL = SQL & ",'" & B1_CEAN & "'"
               SQL = SQL & ",'" & B1_CEANTRIB & "'"
            SQL = SQL & ")"
   
            SQL = Replace(SQL, "]", "")
            SQL = Replace(SQL, "[", "")
   
            CONECTA_GLOBAL.Execute SQL
            Else  'update
               SQL = "update SB1010 set"

                  SQL = SQL & " B1_EX_NCM = '" & B1_EX_NCM & "'"
                  SQL = SQL & ",B1_CODNCM = '" & B1_EX_NCM & "'"
                  SQL = SQL & ",B1_CEAN = '" & B1_CEAN & "'"
                  SQL = SQL & ",B1_CEANTRIB = '" & B1_CEANTRIB & "'"

               SQL = SQL & " where B1_CODANT = '" & Trim(B1_COD) & "'"
               CONECTA_GLOBAL.Execute SQL
         End If
         If TabTempIntegra.State = 1 Then _
            TabTempIntegra.Close
      End If
      TabProdIntegra.MoveNext
   Wend
   If TabProdIntegra.State = 1 Then _
      TabProdIntegra.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INTEGRA_PRODUTO"
End Sub

Sub INTEGRA_FINANCEIRO(E1_PREFIXO As String, ID_N As Long)
'On Error GoTo ERRO_TRATA

   If ID_N <= 0 Then _
      Exit Sub

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabPedidoIntegra As New ADODB.Recordset
   Dim TabCabecaIntegra As New ADODB.Recordset
   Dim TabFinac         As New ADODB.Recordset
   Dim strSQL           As String
   Dim PARCELA_N        As Long
   Dim E1_NOMCLI        As String
   Dim ID_FINAC_N       As Long
   Dim E1_EMISSAO       As String
   Dim E1_VENCTO        As String
   Dim E1_CLIENTE       As String
   Dim E1_CARTAO        As String
   Dim E1_ADM           As String
   Dim E1_CARTAUT       As String
   Dim E1_TIPO          As String
   Dim E1_FILIAL        As String

   A1_FILIAL = "0" & EMPRESA_ID_N
   E1_FILIAL = "0" & ESTABELECIMENTO_ID_N

   If Trim(E1_PREFIXO) = "" Then
      E1_PREFIXO = "NFE"
      Else
         If Trim(E1_PREFIXO) = "55" Then _
            E1_PREFIXO = "NFE"
         If Trim(E1_PREFIXO) = "65" Then _
            E1_PREFIXO = "NFC"
   End If

   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close

   strSQL = "select NF.NF_ID, NF.PESSOA_ID, NF.TRANSP_ID, NF.NF_TIPO, NF.NUMR_NOTA, "
   strSQL = strSQL & " NF.SERIE_NOTA, NF.MODELO_DOC, NF.DT_EMISSAO, NF.STATUS, NF.ESTABELECIMENTO_ID, PESSOA.CNPJCPF, "
   strSQL = strSQL & " PESSOA.DESCRICAO , PESSOA.RAZAO"
   strSQL = strSQL & " from NF WITH (NOLOCK)"
   strSQL = strSQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   strSQL = strSQL & " ON NF.PESSOA_ID = PESSOA.PESSOA_ID"

   strSQL = strSQL & " where nf_id = " & ID_N
   TabPedidoIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoIntegra.EOF Then
      CNPJ_CPF_A = "" & TabPedidoIntegra.Fields("cnpjcpf").Value
      E1_NOMCLI = "" & Trim(Left(TabPedidoIntegra.Fields("descricao").Value, 60))
      Numr_Nota_N = TabPedidoIntegra.Fields("numr_nota").Value
      E1_EMISSAO = "" & TabPedidoIntegra.Fields("dt_emissao").Value
      E1_CLIENTE = "" & TRAZ_ID_TABELA_GLOBAL("SA1010", "A1_COD", "A1_CGC", TabPedidoIntegra.Fields("CNPJCPF").Value)

      If TabFinac.State = 1 Then _
         TabFinac.Close

      SQL = "select ITEMLANCAMENTO.SEQ , ITEMLANCAMENTO.FORMAPAGTO_ID, ITEMLANCAMENTO.Valor_Item, "
      SQL = SQL & " ITEMLANCAMENTO.VALOR_DESCONTO,ITEMLANCAMENTO.NUMR_DP , ITEMLANCAMENTO.DT_VENCIMENTO"
      SQL = SQL & " FROM LANCAMENTO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK)"
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
      SQL = SQL & " where NUMR_DOC = " & ID_N
      SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      TabFinac.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabFinac.EOF
         PARCELA_N = 0 & TabFinac.Fields("seq").Value
         E1_VENCTO = "" & TabFinac.Fields("dt_vencimento").Value
         VALOR_DESCONTO_N = 0 & TabFinac.Fields("valor_desconto").Value
         VALOR_ITEM_N = 0 & (TabFinac.Fields("valor_item").Value - VALOR_DESCONTO_N)
         E1_CARTAO = ""
         E1_ADM = ""
         E1_CARTAUT = ""

'Campo E1_TIPO=Forma de Pagamento gravar os seguintes numeros :
'01=Dinheiro
'02=Cheque
'03=Cartão de Crédito
'04=Cartão de Débito
'05=Crédito Loja
'10=Vale Alimentação
'11=Vale Refeição
'12=Vale Presente
'13=Vale Combustível
'14=Duplicata Mercantil
'90= Sem pagamento
'99=Outros
'onde os numeros informados 03 e 04 e obrigatorios preencher os campos criados no item 03.
'ver relação pois essas formas dependem de cada empresa

         E1_TIPO = "99"

         If TabFinac.Fields("FORMAPAGTO_ID").Value = 1 Then _
            E1_TIPO = "01"
         If TabFinac.Fields("FORMAPAGTO_ID").Value = 2 Then _
            E1_TIPO = "02"

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "select CARTAOPEDIDO_ID,PEDIDO_ID,BANDEIRA_ID,CNPJ_CARTAO,NUMR_AUTORIZACAO from CARTAOPEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & ID_N
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            If Not IsNull(TabConsulta.Fields("CNPJ_CARTAO").Value) Then _
               E1_CARTAO = "" & Trim(TabConsulta.Fields("CNPJ_CARTAO").Value)
            If Not IsNull(TabConsulta.Fields("BANDEIRA_ID").Value) Then
               E1_ADM = "" & Trim(TabConsulta.Fields("BANDEIRA_ID").Value)
               E1_TIPO = "03"
            End If
            If Not IsNull(TabConsulta.Fields("NUMR_AUTORIZACAO").Value) Then _
               E1_CARTAUT = "" & Trim(TabConsulta.Fields("NUMR_AUTORIZACAO").Value)
If Trim(E1_CARTAUT) = "" Then _
   E1_CARTAUT = "000000"
         End If

         ID_FINAC_N = 1

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "select max([R_E_C_N_O_]) from SE1010 WITH (NOLOCK)"
         TabConsulta.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabConsulta.EOF Then _
             If Not IsNull(TabConsulta.Fields(0).Value) Then _
                 ID_FINAC_N = TabConsulta.Fields(0).Value + 1
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         SQL = "select e1_num from SE1010 WITH (NOLOCK)"
         SQL = SQL & " where E1_NUMNOTA = " & Numr_Nota_N
         SQL = SQL & " and e1_parcela = " & PARCELA_N
         SQL = SQL & " and e1_prefixo = '" & E1_PREFIXO & "'"
         TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If TabCabecaIntegra.EOF Then
            strSQL = "insert into SE1010 "
               strSQL = strSQL & "(E1_FILIAL,E1_PREFIXO,E1_NUM,E1_PARCELA,E1_TIPO,E1_NATUREZ"
               strSQL = strSQL & ",E1_PORTADO,E1_AGEDEP,E1_CLIENTE,E1_LOJA,E1_NOMCLI,E1_EMISSAO"
               strSQL = strSQL & ",E1_VENCTO,E1_VENCREA,E1_VALOR,E1_IRRF,E1_ISS,E1_NUMBCO"
               strSQL = strSQL & ",E1_INDICE,E1_BAIXA,E1_NUMBOR,E1_DATABOR,E1_EMIS1,E1_HIST"
               strSQL = strSQL & ",E1_LA,E1_LOTE,E1_MOTIVO,E1_MOVIMEN,E1_OP,E1_SITUACA,E1_CONTRAT"
               strSQL = strSQL & ",E1_SALDO,E1_SUPERVI,E1_VEND1,E1_VEND2,E1_VEND3,E1_VEND4,E1_VEND5"
               strSQL = strSQL & ",E1_COMIS1,E1_COMIS2,E1_COMIS3,E1_COMIS4,E1_DESCONT,E1_COMIS5"
               strSQL = strSQL & ",E1_MULTA,E1_JUROS,E1_CORREC,E1_VALLIQ,E1_VENCORI,E1_CONTA"
               strSQL = strSQL & ",E1_VALJUR,E1_PORCJUR,E1_MOEDA,E1_BASCOM1,E1_BASCOM2,E1_BASCOM3"
               strSQL = strSQL & ",E1_BASCOM4,E1_BASCOM5,E1_FATPREF,E1_FATURA,E1_OK,E1_PROJETO"
               strSQL = strSQL & ",E1_CLASCON,E1_VALCOM1,E1_VALCOM2,E1_VALCOM3,E1_VALCOM4"
               strSQL = strSQL & ",E1_VALCOM5,E1_OCORREN,E1_INSTR1,E1_INSTR2,E1_PEDIDO,E1_DTVARIA"
               strSQL = strSQL & ",E1_VARURV,E1_VLCRUZ,E1_DTFATUR,E1_NUMNOTA,E1_SERIE,E1_STATUS"
               strSQL = strSQL & ",E1_ORIGEM,E1_IDENTEE,E1_NUMCART,E1_FLUXO,E1_DESCFIN,E1_DIADESC"
               strSQL = strSQL & ",E1_CARTAO,E1_CARTVAL,E1_CARTAUT,E1_ADM,E1_VLRREAL,E1_TRANSF"
               strSQL = strSQL & ",E1_BCOCHQ,E1_AGECHQ,E1_CTACHQ,E1_NUMLIQ,E1_ORDPAGO,E1_INSS"
               strSQL = strSQL & ",E1_FILORIG,E1_TIPOFAT,E1_TIPOLIQ,E1_CSLL,E1_COFINS,E1_PIS,E1_FLAGFAT"
               strSQL = strSQL & ",E1_MESBASE,E1_ANOBASE,E1_PLNUCOB,E1_CODINT,E1_CODEMP,E1_MATRIC,E1_ACRESC"
               strSQL = strSQL & ",E1_SDACRES,E1_DECRESC,E1_SDDECRE,E1_MULTNAT,E1_MSFIL,E1_MSEMP"
               strSQL = strSQL & ",E1_PROJPMS,E1_TIPODES,E1_TXMOEDA,E1_DESDOBR,E1_NRDOC,E1_MODSPB"
               strSQL = strSQL & ",E1_IDCNAB,E1_PLCOEMP,E1_PLTPCOE,E1_CODCOR,E1_PARCCSS,D_E_L_E_T_"
               strSQL = strSQL & ",R_E_C_N_O_,E1_DTACRED,E1_NUMCRD,E1_FLAG)"
            strSQL = strSQL & " VALUES( "
               strSQL = strSQL & "'" & E1_FILIAL & "'"               'E1_FILIAL
               strSQL = strSQL & ",'" & E1_PREFIXO & "'"             'E1_PREFIXO
               strSQL = strSQL & ",'" & Numr_Nota_N & "'"            'E1_NUM
               strSQL = strSQL & ",'" & PARCELA_N & "'"              'E1_PARCELA
               strSQL = strSQL & ",'" & E1_TIPO & "'"                'E1_TIPO
               strSQL = strSQL & ",'" & "10101" & "'"                'E1_NATUREZ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PORTADO
               strSQL = strSQL & ",'" & "" & "'"                     'E1_AGEDEP
               strSQL = strSQL & ",'" & E1_CLIENTE & "'"              'E1_CLIENTE
               strSQL = strSQL & ",'" & A1_FILIAL & "'"               'E1_LOJA
               strSQL = strSQL & ",'" & E1_NOMCLI & "'"               'E1_NOMCLI
               strSQL = strSQL & ",'" & E1_EMISSAO & "'"              'E1_EMISSAO
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCTO
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCREA
               strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"   'E1_VALOR
               strSQL = strSQL & ",'" & tpMOEDA(0) & "'"              'E1_IRRF
               strSQL = strSQL & ",'" & tpMOEDA(0) & "'"              'E1_ISS
               strSQL = strSQL & ",'" & "" & "'"                      'E1_NUMBCO
               strSQL = strSQL & ",'" & "" & "'"                      'E1_INDICE
               strSQL = strSQL & ",'" & Null & "'"                    'E1_BAIXA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_NUMBOR
               strSQL = strSQL & ",'" & Null & "'"                    'E1_DATABOR
               strSQL = strSQL & ",'" & Null & "'"                    'E1_EMIS1
               strSQL = strSQL & ",'" & "" & "'"                      'E1_HIST
               strSQL = strSQL & ",'" & "N" & "'"                     'E1_LA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_LOTE
               strSQL = strSQL & ",'" & "" & "'"                      'E1_MOTIVO
               strSQL = strSQL & ",'" & E1_EMISSAO & "'"              'E1_MOVIMEN
               strSQL = strSQL & ",'" & "" & "'"                      'E1_OP
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_SITUACA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_CONTRAT
               strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"   'E1_SALDO
               strSQL = strSQL & ",'" & "" & "'"                      'E1_SUPERVI
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND1
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND2
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND3
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND4
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND5
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS1
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS2
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS3
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS4
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_DESCONT
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS5
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_MULTA
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_JUROS
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_CORREC
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALLIQ
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCORI
               strSQL = strSQL & ",'" & "" & "'"                      'E1_CONTA
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALJUR
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_PORCJUR
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_MOEDA
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM1
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM2
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM3
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM4
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM5
               strSQL = strSQL & ",'" & "" & "'"                      'E1_FATPREF
               strSQL = strSQL & ",'" & "" & "'"                      'E1_FATURA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_OK
               strSQL = strSQL & ",'" & "" & "'"                      'E1_PROJETO
               strSQL = strSQL & ",'" & "" & "'"                      'E1_CLASCON
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM1
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM2
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM3
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM4
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM5
               strSQL = strSQL & ",'" & "01" & "'"                    'E1_OCORREN
               strSQL = strSQL & ",'" & "" & "'"                      'E1_INSTR1
               strSQL = strSQL & ",'" & "" & "'"                      'E1_INSTR2
               strSQL = strSQL & ",'" & "" & "'"                      'E1_PEDIDO
               strSQL = strSQL & ",'" & Null & "'"                   'E1_DTVARIA
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_VARURV
               strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"  'E1_VLCRUZ
               strSQL = strSQL & ",'" & Null & "'"                   'E1_DTFATUR
               strSQL = strSQL & ",'" & Numr_Nota_N & "'"            'E1_NUMNOTA
               strSQL = strSQL & ",'" & "1" & "'"                    'E1_SERIE
               strSQL = strSQL & ",'" & "A" & "'"                    'E1_STATUS
               strSQL = strSQL & ",'" & "" & "'"              'E1_ORIGEM
               strSQL = strSQL & ",'" & "" & "'"                     'E1_IDENTEE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMCART
               strSQL = strSQL & ",'" & "S" & "'"                    'E1_FLUXO
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_DESCFIN
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_DIADESC
               strSQL = strSQL & ",'" & E1_CARTAO & "'"              'E1_CARTAO
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CARTVAL
               strSQL = strSQL & ",'" & E1_CARTAUT & "'"             'E1_CARTAUT
               strSQL = strSQL & ",'" & E1_ADM & "'"                 'E1_ADM
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_VLRREAL
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TRANSF
               strSQL = strSQL & ",'" & "" & "'"                     'E1_BCOCHQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_AGECHQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CTACHQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMLIQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_ORDPAGO
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_INSS
               strSQL = strSQL & ",'" & "" & "'"                     'E1_FILORIG
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPOFAT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPOLIQ
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_CSLL
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_COFINS
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_PIS
               strSQL = strSQL & ",'" & "" & "'"                     'E1_FLAGFAT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MESBASE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_ANOBASE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PLNUCOB
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CODINT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CODEMP
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MATRIC
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_ACRESC
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_SDACRES
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_DECRESC
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_SDDECRE
               strSQL = strSQL & ",'" & "N" & "'"                    'E1_MULTNAT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MSFIL
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MSEMP
               strSQL = strSQL & ",'" & "N" & "'"                    'E1_PROJPMS
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPODES
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_TXMOEDA
               strSQL = strSQL & ",'" & "N" & "'"                    'E1_DESDOBR
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NRDOC
               strSQL = strSQL & ",'" & "1" & "'"                    'E1_MODSPB
               strSQL = strSQL & ",'" & "" & "'"                     'E1_IDCNAB
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PLCOEMP
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PLTPCOE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CODCOR
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PARCCSS
               strSQL = strSQL & ",'" & "" & "'"                     'D_E_L_E_T_
               strSQL = strSQL & "," & ID_FINAC_N                    'R_E_C_N_O_
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_DTACRED
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMCRD
               strSQL = strSQL & ",'" & "" & "'"                     'E1_FLAG
            strSQL = strSQL & ")"

            CONECTA_GLOBAL.Execute strSQL
         End If

         TabFinac.MoveNext
         strSQL = ""
      Wend
      If TabFinac.State = 1 Then _
         TabFinac.Close
   End If
   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INTEGRA_FINANCEIRO"
End Sub
'=============================================
Private Sub Command3_Click()
'On Error GoTo ERRO_TRATA

PEDIDO_ID_N = 0
CONT_N = 0

   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   'ABRE_BANCO_GLOBAL

   'If CONECTA_GLOBAL.State <> 1 Then
      'MsgBox "Banco GLOBAL não conectado."
   '   Exit Sub
   'End If
   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabProdRetaguarda   As New ADODB.Recordset
   Dim TabSB1010           As New ADODB.Recordset
   Dim INDR_GRAVA_PROD     As Boolean

''''''''''''NFE
   If TabSB1010.State = 1 Then _
      TabSB1010.Close

   SQL = "select MFADOC,MFASEQUENCIA,MFAPREFIXO from MFA010 WITH (NOLOCK)"
   SQL = SQL & " where mfadoc not in (select E1_NUMNOTA from SE1010 where E1_PREFIXO = 'NFE' ) "
   SQL = SQL & " AND MFAPREFIXO = 'NFE'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   SQL = SQL & " ORDER BY MFASEQUENCIA"
   TabSB1010.Open SQL, CONECTA_GLOBAL, , , adCmdText
   While Not TabSB1010.EOF

      PEDIDO_ID_N = 0

      If TabProdRetaguarda.State = 1 Then _
         TabProdRetaguarda.Close
      SQL = "select pedido_Id from NF WITH (NOLOCK)"
      SQL = SQL & " where numr_nota = " & TabSB1010.Fields("MFADOC").Value
      SQL = SQL & " and MODELO_DOC = '" & Trim(TabSB1010.Fields("MFAPREFIXO").Value) & "'"
      TabProdRetaguarda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProdRetaguarda.EOF Then _
         If Not IsNull(TabProdRetaguarda.Fields(0).Value) Then _
            PEDIDO_ID_N = 0 & TabProdRetaguarda.Fields(0).Value
      If TabProdRetaguarda.State = 1 Then _
         TabProdRetaguarda.Close

      If PEDIDO_ID_N > 0 Then _
         INTEGRA_FINANCEIRO_acerto TabSB1010.Fields("MFAPREFIXO").Value, TabSB1010.Fields("MFADOC").Value

      CONT_N = CONT_N + 1
      Command3.Caption = CONT_N
      DoEvents

      TabSB1010.MoveNext
   Wend

''''''''''''NFC
   If TabSB1010.State = 1 Then _
      TabSB1010.Close

   SQL = "select MFADOC,MFASEQUENCIA,MFAPREFIXO from MFA010 WITH (NOLOCK)"
   SQL = SQL & " where mfadoc not in (select E1_NUMNOTA from SE1010 where E1_PREFIXO = 'NFC' ) "
   SQL = SQL & " AND MFAPREFIXO = 'NFC'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

   SQL = SQL & " ORDER BY MFASEQUENCIA"
   TabSB1010.Open SQL, CONECTA_GLOBAL, , , adCmdText
   While Not TabSB1010.EOF

      PEDIDO_ID_N = 0

      If TabProdRetaguarda.State = 1 Then _
         TabProdRetaguarda.Close
      SQL = "select pedido_Id from NF WITH (NOLOCK)"
      SQL = SQL & " where numr_nota = " & TabSB1010.Fields("MFADOC").Value
      SQL = SQL & " and MODELO_DOC = '" & Trim(TabSB1010.Fields("MFAPREFIXO").Value) & "'"
      TabProdRetaguarda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProdRetaguarda.EOF Then _
         If Not IsNull(TabProdRetaguarda.Fields(0).Value) Then _
            PEDIDO_ID_N = 0 & TabProdRetaguarda.Fields(0).Value
      If TabProdRetaguarda.State = 1 Then _
         TabProdRetaguarda.Close

      If PEDIDO_ID_N > 0 Then _
         INTEGRA_FINANCEIRO_acerto TabSB1010.Fields("MFAPREFIXO").Value, TabSB1010.Fields("MFADOC").Value

      CONT_N = CONT_N + 1
      Command3.Caption = CONT_N
      DoEvents

      TabSB1010.MoveNext
   Wend
   If TabSB1010.State = 1 Then _
      TabSB1010.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   PEDIDO_ID_N = 0
MsgBox "ok fim"
Exit Sub
ERRO_TRATA:
   PEDIDO_ID_N = 0
   TRATA_ERROS Err.Description, Me.Name, "Command3_Click"
End Sub

Sub INTEGRA_FINANCEIRO_acerto(E1_PREFIXO As String, MFADOC As Long)
'On Error GoTo ERRO_TRATA

   If PEDIDO_ID_N <= 0 Then _
      Exit Sub

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabPedidoIntegra As New ADODB.Recordset
   Dim TabCabecaIntegra As New ADODB.Recordset
   Dim TabFinac         As New ADODB.Recordset
   Dim strSQL           As String
   Dim PARCELA_N        As Long
   Dim E1_NOMCLI        As String
   Dim ID_FINAC_N       As Long
   Dim E1_EMISSAO       As String
   Dim E1_VENCTO        As String
   Dim E1_CLIENTE       As String
   Dim E1_CARTAO        As String
   Dim E1_ADM           As String
   Dim E1_CARTAUT       As String
   Dim E1_TIPO          As String
   Dim E1_FILIAL        As String

   A1_FILIAL = "0" & EMPRESA_ID_N
   E1_FILIAL = "0" & ESTABELECIMENTO_ID_N

   If Trim(E1_PREFIXO) = "" Then
      E1_PREFIXO = "NFE"
      Else
         If Trim(E1_PREFIXO) = "55" Then _
            E1_PREFIXO = "NFE"
         If Trim(E1_PREFIXO) = "65" Then _
            E1_PREFIXO = "NFC"
   End If

   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close

   strSQL = "SELECT * FROM PEDIDO WITH (NOLOCK)"
   strSQL = strSQL & " where pedido_id = " & PEDIDO_ID_N
   strSQL = strSQL & " and status <> 9"
   TabPedidoIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoIntegra.EOF Then
      E1_VENCTO = "" & TabPedidoIntegra.Fields("dt_req").Value
      PESSOA_ID_N = "" & TRAZ_ID_TABELA("PESSOA", "pessoa_id", "cnpjcpf", TabPedidoIntegra.Fields("cgccpf").Value)
      CNPJ_CPF_A = "" & Trim(TabPedidoIntegra.Fields("cgccpf").Value)
      E1_NOMCLI = "" & Trim(Left(TabPedidoIntegra.Fields("nome_cliente").Value, 60))
      Numr_Nota_N = MFADOC
      E1_EMISSAO = "" & TabPedidoIntegra.Fields("dt_REQ").Value
      E1_CLIENTE = "" & TRAZ_ID_TABELA_GLOBAL("SA1010", "A1_COD", "A1_CGC", CNPJ_CPF_A)

      If TabFinac.State = 1 Then _
         TabFinac.Close

      SQL = "SELECT ITEMLANCAMENTO.SEQ, ITEMLANCAMENTO.FORMAPAGTO_ID, ITEMLANCAMENTO.VALOR_ITEM, "
      SQL = SQL & " ITEMLANCAMENTO.VALOR_DESCONTO,ITEMLANCAMENTO.NUMR_DP , ITEMLANCAMENTO.DT_VENCIMENTO"

      SQL = SQL & " FROM ITEMLANCAMENTO WITH (NOLOCK)"

      SQL = SQL & " where NUMR_DOC = " & PEDIDO_ID_N

      TabFinac.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

      If TabFinac.EOF Then
         If TabFinac.State = 1 Then _
            TabFinac.Close

         strSQL = "SELECT sum(qtd_pedida*valor_item) as valor_item FROM PEDIDOitem WITH (NOLOCK)"
         strSQL = strSQL & " where pedido_id = " & PEDIDO_ID_N
         strSQL = strSQL & " and status <> 'C' "
         TabFinac.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
cmdIBGE.Caption = PEDIDO_ID_N
      End If

      While Not TabFinac.EOF
         PARCELA_N = 1
         'E1_VENCTO = "" &
         VALOR_DESCONTO_N = 0 '& TabFinac.Fields("valor_desconto").Value
         VALOR_ITEM_N = 0 & (TabFinac.Fields("valor_item").Value - VALOR_DESCONTO_N)
         E1_CARTAO = ""
         E1_ADM = ""
         E1_CARTAUT = ""

'Campo E1_TIPO=Forma de Pagamento gravar os seguintes numeros :
'01=Dinheiro
'02=Cheque
'03=Cartão de Crédito
'04=Cartão de Débito
'05=Crédito Loja
'10=Vale Alimentação
'11=Vale Refeição
'12=Vale Presente
'13=Vale Combustível
'14=Duplicata Mercantil
'90= Sem pagamento
'99=Outros
'onde os numeros informados 03 e 04 e obrigatorios preencher os campos criados no item 03.
'ver relação pois essas formas dependem de cada empresa

         'If TabFinac.Fields("FORMAPAGTO_ID").Value = 1 Then _
            E1_TIPO = "01"
         'If TabFinac.Fields("FORMAPAGTO_ID").Value = 2 Then _
            E1_TIPO = "02"

         If E1_TIPO = "" Then _
            E1_TIPO = "01"

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "SELECT CARTAOPEDIDO_ID,PEDIDO_ID,BANDEIRA_ID,CNPJ_CARTAO,NUMR_AUTORIZACAO from CARTAOPEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            If Not IsNull(TabConsulta.Fields("CNPJ_CARTAO").Value) Then _
               E1_CARTAO = "" & Trim(TabConsulta.Fields("CNPJ_CARTAO").Value)
            If Not IsNull(TabConsulta.Fields("BANDEIRA_ID").Value) Then
               E1_ADM = "" & Trim(TabConsulta.Fields("BANDEIRA_ID").Value)
               E1_TIPO = "03"
            End If
            If Not IsNull(TabConsulta.Fields("NUMR_AUTORIZACAO").Value) Then _
               E1_CARTAUT = "" & Trim(TabConsulta.Fields("NUMR_AUTORIZACAO").Value)
If Trim(E1_CARTAUT) = "" Then _
   E1_CARTAUT = "000000"
         End If

         ID_FINAC_N = 1

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "select max([R_E_C_N_O_]) from SE1010 WITH (NOLOCK)"
         TabConsulta.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabConsulta.EOF Then _
             If Not IsNull(TabConsulta.Fields(0).Value) Then _
                 ID_FINAC_N = TabConsulta.Fields(0).Value + 1
         If TabConsulta.State = 1 Then _
            TabConsulta.Close

         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         SQL = "select e1_num from SE1010 WITH (NOLOCK)"
         SQL = SQL & " where E1_NUMNOTA = " & Numr_Nota_N
         SQL = SQL & " and e1_parcela = " & PARCELA_N
         SQL = SQL & " and e1_prefixo = '" & E1_PREFIXO & "'"
         TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If TabCabecaIntegra.EOF Then

Command1.Caption = PEDIDO_ID_N

            strSQL = "insert into SE1010 "
               strSQL = strSQL & "(E1_FILIAL,E1_PREFIXO,E1_NUM,E1_PARCELA,E1_TIPO,E1_NATUREZ"
               strSQL = strSQL & ",E1_PORTADO,E1_AGEDEP,E1_CLIENTE,E1_LOJA,E1_NOMCLI,E1_EMISSAO"
               strSQL = strSQL & ",E1_VENCTO,E1_VENCREA,E1_VALOR,E1_IRRF,E1_ISS,E1_NUMBCO"
               strSQL = strSQL & ",E1_INDICE,E1_BAIXA,E1_NUMBOR,E1_DATABOR,E1_EMIS1,E1_HIST"
               strSQL = strSQL & ",E1_LA,E1_LOTE,E1_MOTIVO,E1_MOVIMEN,E1_OP,E1_SITUACA,E1_CONTRAT"
               strSQL = strSQL & ",E1_SALDO,E1_SUPERVI,E1_VEND1,E1_VEND2,E1_VEND3,E1_VEND4,E1_VEND5"
               strSQL = strSQL & ",E1_COMIS1,E1_COMIS2,E1_COMIS3,E1_COMIS4,E1_DESCONT,E1_COMIS5"
               strSQL = strSQL & ",E1_MULTA,E1_JUROS,E1_CORREC,E1_VALLIQ,E1_VENCORI,E1_CONTA"
               strSQL = strSQL & ",E1_VALJUR,E1_PORCJUR,E1_MOEDA,E1_BASCOM1,E1_BASCOM2,E1_BASCOM3"
               strSQL = strSQL & ",E1_BASCOM4,E1_BASCOM5,E1_FATPREF,E1_FATURA,E1_OK,E1_PROJETO"
               strSQL = strSQL & ",E1_CLASCON,E1_VALCOM1,E1_VALCOM2,E1_VALCOM3,E1_VALCOM4"
               strSQL = strSQL & ",E1_VALCOM5,E1_OCORREN,E1_INSTR1,E1_INSTR2,E1_PEDIDO,E1_DTVARIA"
               strSQL = strSQL & ",E1_VARURV,E1_VLCRUZ,E1_DTFATUR,E1_NUMNOTA,E1_SERIE,E1_STATUS"
               strSQL = strSQL & ",E1_ORIGEM,E1_IDENTEE,E1_NUMCART,E1_FLUXO,E1_DESCFIN,E1_DIADESC"
               strSQL = strSQL & ",E1_CARTAO,E1_CARTVAL,E1_CARTAUT,E1_ADM,E1_VLRREAL,E1_TRANSF"
               strSQL = strSQL & ",E1_BCOCHQ,E1_AGECHQ,E1_CTACHQ,E1_NUMLIQ,E1_ORDPAGO,E1_INSS"
               strSQL = strSQL & ",E1_FILORIG,E1_TIPOFAT,E1_TIPOLIQ,E1_CSLL,E1_COFINS,E1_PIS,E1_FLAGFAT"
               strSQL = strSQL & ",E1_MESBASE,E1_ANOBASE,E1_PLNUCOB,E1_CODINT,E1_CODEMP,E1_MATRIC,E1_ACRESC"
               strSQL = strSQL & ",E1_SDACRES,E1_DECRESC,E1_SDDECRE,E1_MULTNAT,E1_MSFIL,E1_MSEMP"
               strSQL = strSQL & ",E1_PROJPMS,E1_TIPODES,E1_TXMOEDA,E1_DESDOBR,E1_NRDOC,E1_MODSPB"
               strSQL = strSQL & ",E1_IDCNAB,E1_PLCOEMP,E1_PLTPCOE,E1_CODCOR,E1_PARCCSS,D_E_L_E_T_"
               strSQL = strSQL & ",R_E_C_N_O_,E1_DTACRED,E1_NUMCRD,E1_FLAG)"
            strSQL = strSQL & " VALUES( "
               strSQL = strSQL & "'" & E1_FILIAL & "'"               'E1_FILIAL
               strSQL = strSQL & ",'" & E1_PREFIXO & "'"             'E1_PREFIXO
               strSQL = strSQL & ",'" & Numr_Nota_N & "'"            'E1_NUM
               strSQL = strSQL & ",'" & PARCELA_N & "'"              'E1_PARCELA
               strSQL = strSQL & ",'" & E1_TIPO & "'"                'E1_TIPO
               strSQL = strSQL & ",'" & "10101" & "'"                'E1_NATUREZ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PORTADO
               strSQL = strSQL & ",'" & "" & "'"                     'E1_AGEDEP
               strSQL = strSQL & ",'" & E1_CLIENTE & "'"              'E1_CLIENTE
               strSQL = strSQL & ",'" & A1_FILIAL & "'"               'E1_LOJA
               strSQL = strSQL & ",'" & E1_NOMCLI & "'"               'E1_NOMCLI
               strSQL = strSQL & ",'" & E1_EMISSAO & "'"              'E1_EMISSAO
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCTO
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCREA
               strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"   'E1_VALOR
               strSQL = strSQL & ",'" & tpMOEDA(0) & "'"              'E1_IRRF
               strSQL = strSQL & ",'" & tpMOEDA(0) & "'"              'E1_ISS
               strSQL = strSQL & ",'" & "" & "'"                      'E1_NUMBCO
               strSQL = strSQL & ",'" & "" & "'"                      'E1_INDICE
               strSQL = strSQL & ",'" & Null & "'"                    'E1_BAIXA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_NUMBOR
               strSQL = strSQL & ",'" & Null & "'"                    'E1_DATABOR
               strSQL = strSQL & ",'" & Null & "'"                    'E1_EMIS1
               strSQL = strSQL & ",'" & "" & "'"                      'E1_HIST
               strSQL = strSQL & ",'" & "N" & "'"                     'E1_LA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_LOTE
               strSQL = strSQL & ",'" & "" & "'"                      'E1_MOTIVO
               strSQL = strSQL & ",'" & E1_EMISSAO & "'"              'E1_MOVIMEN
               strSQL = strSQL & ",'" & "" & "'"                      'E1_OP
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_SITUACA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_CONTRAT
               strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"   'E1_SALDO
               strSQL = strSQL & ",'" & "" & "'"                      'E1_SUPERVI
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND1
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND2
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND3
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND4
               strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND5
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS1
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS2
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS3
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS4
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_DESCONT
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS5
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_MULTA
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_JUROS
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_CORREC
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALLIQ
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCORI
               strSQL = strSQL & ",'" & "" & "'"                      'E1_CONTA
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALJUR
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_PORCJUR
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_MOEDA
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM1
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM2
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM3
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM4
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM5
               strSQL = strSQL & ",'" & "" & "'"                      'E1_FATPREF
               strSQL = strSQL & ",'" & "" & "'"                      'E1_FATURA
               strSQL = strSQL & ",'" & "" & "'"                      'E1_OK
               strSQL = strSQL & ",'" & "" & "'"                      'E1_PROJETO
               strSQL = strSQL & ",'" & "" & "'"                      'E1_CLASCON
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM1
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM2
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM3
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM4
               strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM5
               strSQL = strSQL & ",'" & "01" & "'"                    'E1_OCORREN
               strSQL = strSQL & ",'" & "" & "'"                      'E1_INSTR1
               strSQL = strSQL & ",'" & "" & "'"                      'E1_INSTR2
               strSQL = strSQL & ",'" & "" & "'"                      'E1_PEDIDO
               strSQL = strSQL & ",'" & Null & "'"                   'E1_DTVARIA
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_VARURV
               strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"  'E1_VLCRUZ
               strSQL = strSQL & ",'" & Null & "'"                   'E1_DTFATUR
               strSQL = strSQL & ",'" & Numr_Nota_N & "'"            'E1_NUMNOTA
               strSQL = strSQL & ",'" & "1" & "'"                    'E1_SERIE
               strSQL = strSQL & ",'" & "A" & "'"                    'E1_STATUS
               strSQL = strSQL & ",'" & "" & "'"              'E1_ORIGEM
               strSQL = strSQL & ",'" & "" & "'"                     'E1_IDENTEE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMCART
               strSQL = strSQL & ",'" & "S" & "'"                    'E1_FLUXO
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_DESCFIN
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_DIADESC
               strSQL = strSQL & ",'" & E1_CARTAO & "'"              'E1_CARTAO
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CARTVAL
               strSQL = strSQL & ",'" & E1_CARTAUT & "'"             'E1_CARTAUT
               strSQL = strSQL & ",'" & E1_ADM & "'"                 'E1_ADM
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_VLRREAL
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TRANSF
               strSQL = strSQL & ",'" & "" & "'"                     'E1_BCOCHQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_AGECHQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CTACHQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMLIQ
               strSQL = strSQL & ",'" & "" & "'"                     'E1_ORDPAGO
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_INSS
               strSQL = strSQL & ",'" & "" & "'"                     'E1_FILORIG
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPOFAT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPOLIQ
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_CSLL
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_COFINS
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_PIS
               strSQL = strSQL & ",'" & "" & "'"                     'E1_FLAGFAT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MESBASE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_ANOBASE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PLNUCOB
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CODINT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CODEMP
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MATRIC
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_ACRESC
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_SDACRES
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_DECRESC
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_SDDECRE
               strSQL = strSQL & ",'" & "N" & "'"                    'E1_MULTNAT
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MSFIL
               strSQL = strSQL & ",'" & "" & "'"                     'E1_MSEMP
               strSQL = strSQL & ",'" & "N" & "'"                    'E1_PROJPMS
               strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPODES
               strSQL = strSQL & ",'" & "0" & "'"                    'E1_TXMOEDA
               strSQL = strSQL & ",'" & "N" & "'"                    'E1_DESDOBR
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NRDOC
               strSQL = strSQL & ",'" & "1" & "'"                    'E1_MODSPB
               strSQL = strSQL & ",'" & "" & "'"                     'E1_IDCNAB
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PLCOEMP
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PLTPCOE
               strSQL = strSQL & ",'" & "" & "'"                     'E1_CODCOR
               strSQL = strSQL & ",'" & "" & "'"                     'E1_PARCCSS
               strSQL = strSQL & ",'" & "" & "'"                     'D_E_L_E_T_
               strSQL = strSQL & "," & ID_FINAC_N                    'R_E_C_N_O_
               strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_DTACRED
               strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMCRD
               strSQL = strSQL & ",'" & "" & "'"                     'E1_FLAG
            strSQL = strSQL & ")"

            CONECTA_GLOBAL.Execute strSQL
         End If

         TabFinac.MoveNext
         strSQL = ""
      Wend
      If TabFinac.State = 1 Then _
         TabFinac.Close
   End If
   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close
   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INTEGRA_FINANCEIRO_acerto"
End Sub

Sub INTEGRA_FINANCEIRO_UM_SO(E1_PREFIXO As String, MFADOC As Long, CNPJ_CPF_A As String)
'On Error GoTo ERRO_TRATA

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabPedidoIntegra As New ADODB.Recordset
   Dim TabCabecaIntegra As New ADODB.Recordset
   Dim TabFinac         As New ADODB.Recordset
   Dim strSQL           As String
   Dim PARCELA_N        As Long
   Dim E1_NOMCLI        As String
   Dim ID_FINAC_N       As Long
   Dim E1_EMISSAO       As String
   Dim E1_VENCTO        As String
   Dim E1_CLIENTE       As String
   Dim E1_CARTAO        As String
   Dim E1_ADM           As String
   Dim E1_CARTAUT       As String
   Dim E1_TIPO          As String
   Dim E1_FILIAL        As String

   A1_FILIAL = "0" & EMPRESA_ID_N
   E1_FILIAL = "0" & ESTABELECIMENTO_ID_N

   If Trim(E1_PREFIXO) = "" Then
      E1_PREFIXO = "NFE"
      Else
         If Trim(E1_PREFIXO) = "55" Then _
            E1_PREFIXO = "NFE"
         If Trim(E1_PREFIXO) = "65" Then _
            E1_PREFIXO = "NFC"
   End If

      E1_VENCTO = "" & DMA(Date)
      'PESSOA_ID_N = "" & TRAZ_ID_TABELA("PESSOA", "pessoa_id", "cnpjcpf", TabPedidoIntegra.Fields("cgccpf").Value)
      E1_NOMCLI = "Consumidor Final"
      Numr_Nota_N = MFADOC
      E1_EMISSAO = "" & DMA(Date)
      E1_CLIENTE = "" & TRAZ_ID_TABELA_GLOBAL("SA1010", "A1_COD", "A1_CGC", CNPJ_CPF_A)

      PARCELA_N = 1
      VALOR_DESCONTO_N = 0
      VALOR_ITEM_N = 1
      E1_CARTAO = ""
      E1_ADM = ""
      E1_CARTAUT = ""
      E1_TIPO = "01"
      ID_FINAC_N = 1

      SQL = "delete SE1010 "
      SQL = SQL & " where E1_NUMNOTA = " & Numr_Nota_N
      SQL = SQL & " and e1_parcela = " & PARCELA_N
      SQL = SQL & " and e1_prefixo = '" & E1_PREFIXO & "'"
      CONECTA_GLOBAL.Execute SQL

      If TabConsulta.State = 1 Then _
         TabConsulta.Close
      SQL = "select max([R_E_C_N_O_]) from SE1010 WITH (NOLOCK)"
      TabConsulta.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabConsulta.EOF Then _
          If Not IsNull(TabConsulta.Fields(0).Value) Then _
              ID_FINAC_N = TabConsulta.Fields(0).Value + 1
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close
      SQL = "select e1_num from SE1010 WITH (NOLOCK)"
      SQL = SQL & " where E1_NUMNOTA = " & Numr_Nota_N
      SQL = SQL & " and e1_parcela = " & PARCELA_N
      SQL = SQL & " and e1_prefixo = '" & E1_PREFIXO & "'"
      TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If TabCabecaIntegra.EOF Then
         strSQL = "insert into SE1010 "
            strSQL = strSQL & "(E1_FILIAL,E1_PREFIXO,E1_NUM,E1_PARCELA,E1_TIPO,E1_NATUREZ"
            strSQL = strSQL & ",E1_PORTADO,E1_AGEDEP,E1_CLIENTE,E1_LOJA,E1_NOMCLI,E1_EMISSAO"
            strSQL = strSQL & ",E1_VENCTO,E1_VENCREA,E1_VALOR,E1_IRRF,E1_ISS,E1_NUMBCO"
            strSQL = strSQL & ",E1_INDICE,E1_BAIXA,E1_NUMBOR,E1_DATABOR,E1_EMIS1,E1_HIST"
            strSQL = strSQL & ",E1_LA,E1_LOTE,E1_MOTIVO,E1_MOVIMEN,E1_OP,E1_SITUACA,E1_CONTRAT"
            strSQL = strSQL & ",E1_SALDO,E1_SUPERVI,E1_VEND1,E1_VEND2,E1_VEND3,E1_VEND4,E1_VEND5"
            strSQL = strSQL & ",E1_COMIS1,E1_COMIS2,E1_COMIS3,E1_COMIS4,E1_DESCONT,E1_COMIS5"
            strSQL = strSQL & ",E1_MULTA,E1_JUROS,E1_CORREC,E1_VALLIQ,E1_VENCORI,E1_CONTA"
            strSQL = strSQL & ",E1_VALJUR,E1_PORCJUR,E1_MOEDA,E1_BASCOM1,E1_BASCOM2,E1_BASCOM3"
            strSQL = strSQL & ",E1_BASCOM4,E1_BASCOM5,E1_FATPREF,E1_FATURA,E1_OK,E1_PROJETO"
            strSQL = strSQL & ",E1_CLASCON,E1_VALCOM1,E1_VALCOM2,E1_VALCOM3,E1_VALCOM4"
            strSQL = strSQL & ",E1_VALCOM5,E1_OCORREN,E1_INSTR1,E1_INSTR2,E1_PEDIDO,E1_DTVARIA"
            strSQL = strSQL & ",E1_VARURV,E1_VLCRUZ,E1_DTFATUR,E1_NUMNOTA,E1_SERIE,E1_STATUS"
            strSQL = strSQL & ",E1_ORIGEM,E1_IDENTEE,E1_NUMCART,E1_FLUXO,E1_DESCFIN,E1_DIADESC"
            strSQL = strSQL & ",E1_CARTAO,E1_CARTVAL,E1_CARTAUT,E1_ADM,E1_VLRREAL,E1_TRANSF"
            strSQL = strSQL & ",E1_BCOCHQ,E1_AGECHQ,E1_CTACHQ,E1_NUMLIQ,E1_ORDPAGO,E1_INSS"
            strSQL = strSQL & ",E1_FILORIG,E1_TIPOFAT,E1_TIPOLIQ,E1_CSLL,E1_COFINS,E1_PIS,E1_FLAGFAT"
            strSQL = strSQL & ",E1_MESBASE,E1_ANOBASE,E1_PLNUCOB,E1_CODINT,E1_CODEMP,E1_MATRIC,E1_ACRESC"
            strSQL = strSQL & ",E1_SDACRES,E1_DECRESC,E1_SDDECRE,E1_MULTNAT,E1_MSFIL,E1_MSEMP"
            strSQL = strSQL & ",E1_PROJPMS,E1_TIPODES,E1_TXMOEDA,E1_DESDOBR,E1_NRDOC,E1_MODSPB"
            strSQL = strSQL & ",E1_IDCNAB,E1_PLCOEMP,E1_PLTPCOE,E1_CODCOR,E1_PARCCSS,D_E_L_E_T_"
            strSQL = strSQL & ",R_E_C_N_O_,E1_DTACRED,E1_NUMCRD,E1_FLAG)"
         strSQL = strSQL & " VALUES( "
            strSQL = strSQL & "'" & E1_FILIAL & "'"               'E1_FILIAL
            strSQL = strSQL & ",'" & E1_PREFIXO & "'"             'E1_PREFIXO
            strSQL = strSQL & ",'" & Numr_Nota_N & "'"            'E1_NUM
            strSQL = strSQL & ",'" & PARCELA_N & "'"              'E1_PARCELA
            strSQL = strSQL & ",'" & E1_TIPO & "'"                'E1_TIPO
            strSQL = strSQL & ",'" & "10101" & "'"                'E1_NATUREZ
            strSQL = strSQL & ",'" & "" & "'"                     'E1_PORTADO
            strSQL = strSQL & ",'" & "" & "'"                     'E1_AGEDEP
            strSQL = strSQL & ",'" & E1_CLIENTE & "'"              'E1_CLIENTE
            strSQL = strSQL & ",'" & A1_FILIAL & "'"               'E1_LOJA
            strSQL = strSQL & ",'" & E1_NOMCLI & "'"               'E1_NOMCLI
            strSQL = strSQL & ",'" & E1_EMISSAO & "'"              'E1_EMISSAO
            strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCTO
            strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCREA
            strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"   'E1_VALOR
            strSQL = strSQL & ",'" & tpMOEDA(0) & "'"              'E1_IRRF
            strSQL = strSQL & ",'" & tpMOEDA(0) & "'"              'E1_ISS
            strSQL = strSQL & ",'" & "" & "'"                      'E1_NUMBCO
            strSQL = strSQL & ",'" & "" & "'"                      'E1_INDICE
            strSQL = strSQL & ",'" & Null & "'"                    'E1_BAIXA
            strSQL = strSQL & ",'" & "" & "'"                      'E1_NUMBOR
            strSQL = strSQL & ",'" & Null & "'"                    'E1_DATABOR
            strSQL = strSQL & ",'" & Null & "'"                    'E1_EMIS1
            strSQL = strSQL & ",'" & "" & "'"                      'E1_HIST
            strSQL = strSQL & ",'" & "N" & "'"                     'E1_LA
            strSQL = strSQL & ",'" & "" & "'"                      'E1_LOTE
            strSQL = strSQL & ",'" & "" & "'"                      'E1_MOTIVO
            strSQL = strSQL & ",'" & E1_EMISSAO & "'"              'E1_MOVIMEN
            strSQL = strSQL & ",'" & "" & "'"                      'E1_OP
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_SITUACA
            strSQL = strSQL & ",'" & "" & "'"                      'E1_CONTRAT
            strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"   'E1_SALDO
            strSQL = strSQL & ",'" & "" & "'"                      'E1_SUPERVI
            strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND1
            strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND2
            strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND3
            strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND4
            strSQL = strSQL & ",'" & "" & "'"                      'E1_VEND5
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS1
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS2
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS3
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS4
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_DESCONT
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_COMIS5
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_MULTA
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_JUROS
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_CORREC
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALLIQ
            strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_VENCORI
            strSQL = strSQL & ",'" & "" & "'"                      'E1_CONTA
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALJUR
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_PORCJUR
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_MOEDA
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM1
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM2
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM3
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM4
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_BASCOM5
            strSQL = strSQL & ",'" & "" & "'"                      'E1_FATPREF
            strSQL = strSQL & ",'" & "" & "'"                      'E1_FATURA
            strSQL = strSQL & ",'" & "" & "'"                      'E1_OK
            strSQL = strSQL & ",'" & "" & "'"                      'E1_PROJETO
            strSQL = strSQL & ",'" & "" & "'"                      'E1_CLASCON
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM1
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM2
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM3
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM4
            strSQL = strSQL & ",'" & "0" & "'"                     'E1_VALCOM5
            strSQL = strSQL & ",'" & "01" & "'"                    'E1_OCORREN
            strSQL = strSQL & ",'" & "" & "'"                      'E1_INSTR1
            strSQL = strSQL & ",'" & "" & "'"                      'E1_INSTR2
            strSQL = strSQL & ",'" & "" & "'"                      'E1_PEDIDO
            strSQL = strSQL & ",'" & Null & "'"                   'E1_DTVARIA
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_VARURV
            strSQL = strSQL & ",'" & tpMOEDA(VALOR_ITEM_N) & "'"  'E1_VLCRUZ
            strSQL = strSQL & ",'" & Null & "'"                   'E1_DTFATUR
            strSQL = strSQL & ",'" & Numr_Nota_N & "'"            'E1_NUMNOTA
            strSQL = strSQL & ",'" & "1" & "'"                    'E1_SERIE
            strSQL = strSQL & ",'" & "A" & "'"                    'E1_STATUS
            strSQL = strSQL & ",'" & "" & "'"              'E1_ORIGEM
            strSQL = strSQL & ",'" & "" & "'"                     'E1_IDENTEE
            strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMCART
            strSQL = strSQL & ",'" & "S" & "'"                    'E1_FLUXO
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_DESCFIN
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_DIADESC
            strSQL = strSQL & ",'" & E1_CARTAO & "'"              'E1_CARTAO
            strSQL = strSQL & ",'" & "" & "'"                     'E1_CARTVAL
            strSQL = strSQL & ",'" & E1_CARTAUT & "'"             'E1_CARTAUT
            strSQL = strSQL & ",'" & E1_ADM & "'"                 'E1_ADM
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_VLRREAL
            strSQL = strSQL & ",'" & "" & "'"                     'E1_TRANSF
            strSQL = strSQL & ",'" & "" & "'"                     'E1_BCOCHQ
            strSQL = strSQL & ",'" & "" & "'"                     'E1_AGECHQ
            strSQL = strSQL & ",'" & "" & "'"                     'E1_CTACHQ
            strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMLIQ
            strSQL = strSQL & ",'" & "" & "'"                     'E1_ORDPAGO
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_INSS
            strSQL = strSQL & ",'" & "" & "'"                     'E1_FILORIG
            strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPOFAT
            strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPOLIQ
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_CSLL
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_COFINS
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_PIS
            strSQL = strSQL & ",'" & "" & "'"                     'E1_FLAGFAT
            strSQL = strSQL & ",'" & "" & "'"                     'E1_MESBASE
            strSQL = strSQL & ",'" & "" & "'"                     'E1_ANOBASE
            strSQL = strSQL & ",'" & "" & "'"                     'E1_PLNUCOB
            strSQL = strSQL & ",'" & "" & "'"                     'E1_CODINT
            strSQL = strSQL & ",'" & "" & "'"                     'E1_CODEMP
            strSQL = strSQL & ",'" & "" & "'"                     'E1_MATRIC
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_ACRESC
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_SDACRES
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_DECRESC
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_SDDECRE
            strSQL = strSQL & ",'" & "N" & "'"                    'E1_MULTNAT
            strSQL = strSQL & ",'" & "" & "'"                     'E1_MSFIL
            strSQL = strSQL & ",'" & "" & "'"                     'E1_MSEMP
            strSQL = strSQL & ",'" & "N" & "'"                    'E1_PROJPMS
            strSQL = strSQL & ",'" & "" & "'"                     'E1_TIPODES
            strSQL = strSQL & ",'" & "0" & "'"                    'E1_TXMOEDA
            strSQL = strSQL & ",'" & "N" & "'"                    'E1_DESDOBR
            strSQL = strSQL & ",'" & "" & "'"                     'E1_NRDOC
            strSQL = strSQL & ",'" & "1" & "'"                    'E1_MODSPB
            strSQL = strSQL & ",'" & "" & "'"                     'E1_IDCNAB
            strSQL = strSQL & ",'" & "" & "'"                     'E1_PLCOEMP
            strSQL = strSQL & ",'" & "" & "'"                     'E1_PLTPCOE
            strSQL = strSQL & ",'" & "" & "'"                     'E1_CODCOR
            strSQL = strSQL & ",'" & "" & "'"                     'E1_PARCCSS
            strSQL = strSQL & ",'" & "" & "'"                     'D_E_L_E_T_
            strSQL = strSQL & "," & ID_FINAC_N                    'R_E_C_N_O_
            strSQL = strSQL & ",'" & E1_VENCTO & "'"               'E1_DTACRED
            strSQL = strSQL & ",'" & "" & "'"                     'E1_NUMCRD
            strSQL = strSQL & ",'" & "" & "'"                     'E1_FLAG
         strSQL = strSQL & ")"

         CONECTA_GLOBAL.Execute strSQL
      End If
      strSQL = ""

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "INTEGRA_FINANCEIRO_UM_SO"
End Sub

Sub CLIENTE_INTEGRA(CNPJ_CPF_A As String)
'On Error GoTo ERRO_TRATA

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabCliIntegra    As New ADODB.Recordset
   Dim TabTempIntegra   As New ADODB.Recordset
   Dim TabEndIntegra    As New ADODB.Recordset
   Dim TabFoneIntegra   As New ADODB.Recordset
   Dim strSQL           As String
   Dim A1_NUMERO        As String
   Dim A1_CODCIDADE     As String
   Dim A1_ENENTCGC      As String
   Dim A1_ENDENTNR      As String
   Dim A1_CEPE          As String
   Dim A1_MUNE          As String
   Dim A1_ESTE          As String
   Dim A1_CODCIDENT     As String
   Dim A1_UFENTREGA     As String
   Dim A1_SUFRAMA       As String
   Dim A1_COD           As Long

   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close

   strSQL = "select * from PESSOA WITH (NOLOCK)"
   strSQL = strSQL & " WHERE len(cnpjcpf) >= 11 "
   strSQL = strSQL & " and descricao <> '' "

   If Trim(CNPJ_CPF_A) <> "" Then
      strSQL = strSQL & " and cnpjcpf = '" & Trim(CNPJ_CPF_A) & "'"
      Else: strSQL = strSQL & " ORDER BY descricao "
   End If

   TabCliIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText

   While Not TabCliIntegra.EOF
      LIMPA_VARIAVEIS

      PESSOA_ID_N = TabCliIntegra.Fields("pessoa_id").Value

      A1_FILIAL = "0" & EMPRESA_ID_N
      A1_LOJA = "0" & ESTABELECIMENTO_ID_N
      A1_NOME = "" & Left(Trim(TabCliIntegra.Fields("DESCRICAO").Value), 70)
      A1_NOME = Trim(Replace(A1_NOME, ",", "."))
      A1_CGC = "" & Trim(TabCliIntegra.Fields("cnpjcpf").Value)
      A1_CONTATO = "" '& Trim(TabCliIntegra.Fields("contato").Value)

      If Len(Trim(TabCliIntegra.Fields("CNPJcpf").Value)) <= 11 Then
         A1_PESSOA = "F"
         Else: A1_PESSOA = "J"
      End If

      A1_NREDUZ = "" & Left(Trim(A1_NOME), 45)
      A1_TIPO = "F"

      A1_INSCR = "" & Trim(TRAZ_IE(PESSOA_ID_N))
      If Trim(A1_INSCR) = "" Or Trim(A1_INSCR) = "0" Then _
         A1_INSCR = "ISENTO"

      A1_INSCRM = "" & Trim(TRAZ_IM(PESSOA_ID_N))
      'If Trim(A1_INSCRM) = "" Then _
         A1_INSCRM = "ISENTO"

      A1_PFISICA = ""
      A1_RG = "" & Trim(TRAZ_RG(PESSOA_ID_N))
      A1_EMAIL = "" & Trim(TRAZ_EMAIL(PESSOA_ID_N))
      A1_HPAGE = ""
      A1_INSCRUR = ""

'==========ENDEREÇO
      ENDERECO_A = ""
      A1_END = ""
      A1_NUMERO = ""
      A1_MUN = ""
      A1_MUNE = ""
      A1_EST = ""
      A1_UFENTREGA = ""
      A1_ESTE = ""
      A1_BAIRRO = ""
      A1_ESTADO = ""
      A1_CEP = ""
      A1_CEPE = ""
      A1_CODCIDADE = ""
      A1_CODCIDENT = ""
      A1_ENDENTNR = ""
      A1_ENDENTNR = "S/N"

      If TabEndIntegra.State = 1 Then _
         TabEndIntegra.Close

      strSQL = "select ENDERECO.ENDERECO_ID, ENDERECO.CEP_ID, ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, ENDERECO.TIPO, ENDERECO.NUMERO, "
      strSQL = strSQL & " CLIENTE.CLIENTE_ID, CLIENTE.ESTABELECIMENTO_ID, CLIENTE.VENDEDOR_ID, CLIENTE.CGCCPF, CLIENTE.NOME, CLIENTE.RAZAO_SOCIAL, CLIENTE.STATUS,"
      strSQL = strSQL & " CEP.Cidade , CEP.UF, CEP.IBGE_ID"
      strSQL = strSQL & " from CLIENTE WITH (NOLOCK)"
      strSQL = strSQL & " LEFT OUTER JOIN CEP WITH (NOLOCK)"
      strSQL = strSQL & " RIGHT OUTER JOIN ENDERECO WITH (NOLOCK)"
      strSQL = strSQL & " ON CEP.CEP_ID = ENDERECO.CEP_ID "
      strSQL = strSQL & " ON CLIENTE.PESSOA_ID = ENDERECO.PESSOA_ID"

strSQL = strSQL & " where CLIENTE.pessoa_id = " & PESSOA_ID_N
strSQL = strSQL & " and tipo = 'C'"

      TabEndIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEndIntegra.EOF Then
         ENDERECO_A = ""
         ENDERECO_A = "" & TabEndIntegra.Fields("rua").Value & " " & TabEndIntegra.Fields("complemento").Value
         ENDERECO_A = Trim(Replace(ENDERECO_A, ",", "."))
         ENDERECO_A = Left(Trim(ENDERECO_A), 60)

         A1_END = "" & ENDERECO_A
         A1_NUMERO = "" & Trim(TabEndIntegra.Fields("numero").Value)
         A1_MUN = "" & Trim(TabEndIntegra.Fields("cidade").Value)
         A1_MUNE = "" & Trim(TabEndIntegra.Fields("cidade").Value)
         A1_EST = "" & Trim(TabEndIntegra.Fields("uf").Value)
         A1_UFENTREGA = "" & Trim(TabEndIntegra.Fields("uf").Value)
         A1_ESTE = "" & Trim(TabEndIntegra.Fields("uf").Value)
         A1_BAIRRO = "" & Trim(TabEndIntegra.Fields("bairro").Value)
         A1_ESTADO = "" & Trim(TabEndIntegra.Fields("uf").Value)
         A1_CEP = "" & Trim(TabEndIntegra.Fields("cep_id").Value)
         A1_CEPE = "" & Trim(TabEndIntegra.Fields("cep_id").Value)
         A1_CODCIDADE = "" & Trim(TabEndIntegra.Fields("IBGE_ID").Value)
         A1_CODCIDENT = "" & Trim(TabEndIntegra.Fields("IBGE_ID").Value)
         A1_ENDENTNR = "" & A1_NUMERO
         If Trim(A1_ENDENTNR) = "" Then _
            A1_ENDENTNR = "S/N"
         Else
            'se não achou vai integrar endereço da empresas
            If TabEndIntegra.State = 1 Then _
               TabEndIntegra.Close
      
            strSQL = "select ENDERECO.ENDERECO_ID, ENDERECO.CEP_ID, ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, ENDERECO.TIPO, ENDERECO.NUMERO, "
            strSQL = strSQL & " CLIENTE.CLIENTE_ID, CLIENTE.ESTABELECIMENTO_ID, CLIENTE.VENDEDOR_ID, CLIENTE.CGCCPF, CLIENTE.NOME, CLIENTE.RAZAO_SOCIAL, CLIENTE.STATUS,"
            strSQL = strSQL & " CEP.Cidade , CEP.UF, CEP.IBGE_ID"
            strSQL = strSQL & " from CLIENTE WITH (NOLOCK)"
            strSQL = strSQL & " LEFT OUTER JOIN CEP WITH (NOLOCK)"
            strSQL = strSQL & " RIGHT OUTER JOIN ENDERECO WITH (NOLOCK)"
            strSQL = strSQL & " ON CEP.CEP_ID = ENDERECO.CEP_ID "
            strSQL = strSQL & " ON CLIENTE.PESSOA_ID = ENDERECO.PESSOA_ID"
      
            strSQL = strSQL & " where CLIENTE.pessoa_id = 1"
            strSQL = strSQL & " and tipo = 'C'"
      
            TabEndIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabEndIntegra.EOF Then
               ENDERECO_A = ""
               ENDERECO_A = "" & TabEndIntegra.Fields("rua").Value & " " & TabEndIntegra.Fields("complemento").Value
               ENDERECO_A = Trim(Replace(ENDERECO_A, ",", "."))
               ENDERECO_A = Left(Trim(ENDERECO_A), 60)
               A1_END = "" & ENDERECO_A
               A1_NUMERO = "" & Trim(TabEndIntegra.Fields("numero").Value)
               A1_MUN = "" & Trim(TabEndIntegra.Fields("cidade").Value)
               A1_MUNE = "" & Trim(TabEndIntegra.Fields("cidade").Value)
               A1_EST = "" & Trim(TabEndIntegra.Fields("uf").Value)
               A1_UFENTREGA = "" & Trim(TabEndIntegra.Fields("uf").Value)
               A1_ESTE = "" & Trim(TabEndIntegra.Fields("uf").Value)
               A1_BAIRRO = "" & Trim(TabEndIntegra.Fields("bairro").Value)
               A1_ESTADO = "" & Trim(TabEndIntegra.Fields("uf").Value)
               A1_CEP = "" & Trim(TabEndIntegra.Fields("cep_id").Value)
               A1_CEPE = "" & Trim(TabEndIntegra.Fields("cep_id").Value)
               A1_CODCIDADE = "" & Trim(TabEndIntegra.Fields("IBGE_ID").Value)
               A1_CODCIDENT = "" & Trim(TabEndIntegra.Fields("IBGE_ID").Value)
               A1_ENDENTNR = "" & A1_NUMERO
               If Trim(A1_ENDENTNR) = "" Then _
                  A1_ENDENTNR = "S/N"
            End If
      End If
      If TabEndIntegra.State = 1 Then _
         TabEndIntegra.Close

'==========FONE
      A1_DDI = ""
      A1_DDD = ""
      A1_TEL = ""
      A1_TELEX = ""
      A1_FAX = ""

      If TabFoneIntegra.State = 1 Then _
         TabFoneIntegra.Close

      strSQL = "select * from FONE WITH (NOLOCK)"
      strSQL = strSQL & " where pessoa_id = " & PESSOA_ID_N
      TabFoneIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFoneIntegra.EOF Then
         A1_DDI = "" & Trim(TabFoneIntegra.Fields("ddd").Value)
         A1_DDD = "" & Trim(TabFoneIntegra.Fields("ddd").Value)
         A1_TEL = "" & Trim(TabFoneIntegra.Fields("numero").Value)
         A1_TELEX = "" & Trim(TabFoneIntegra.Fields("numero").Value)
         A1_FAX = "" & Trim(TabFoneIntegra.Fields("numero").Value)
      End If
      If TabFoneIntegra.State = 1 Then _
         TabFoneIntegra.Close

      If TabTempIntegra.State = 1 Then _
         TabTempIntegra.Close
      strSQL = "select A1_COD from SA1010 WITH (NOLOCK)"
      strSQL = strSQL & " where A1_CGC = '" & Trim(A1_CGC) & "'"
      TabTempIntegra.Open strSQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabTempIntegra.EOF Then
         A1_COD = 0 & TabTempIntegra.Fields("A1_COD").Value
         Acao_N = 2
         Else
            Acao_N = 1

            If TabTempIntegra.State = 1 Then _
               TabTempIntegra.Close
            strSQL = "select max(CONVERT(INT,a1_cod)) from sa1010 WITH (NOLOCK)"
            TabTempIntegra.Open strSQL, CONECTA_GLOBAL, , , adCmdText
            If Not TabTempIntegra.EOF Then _
               If Not IsNull(TabTempIntegra.Fields(0).Value) Then _
                  A1_COD = 1 + TabTempIntegra.Fields(0).Value
      End If
      If TabTempIntegra.State = 1 Then _
         TabTempIntegra.Close

      If Trim(A1_CODCIDADE) = "" Or Trim(A1_CEP) = "" Or Trim(A1_MUN) = "" Or Trim(A1_EST) = "" Then
         MsgBox "Erro no cadastro endereço, verificar. Cep = " & A1_CEP
         Else
            SQL = "spClienteGlobal " & Acao_N _
                                     & ",'" & A1_FILIAL & "'," & A1_COD & ",'" & A1_LOJA & "','" & A1_NOME _
                                     & "','" & A1_PESSOA & "','" & A1_NREDUZ & "','" & A1_TIPO & "','" & A1_END _
                                     & "','" & A1_NUMERO & "','" & A1_MUN & "','" & A1_EST & "','" & A1_BAIRRO _
                                     & "','" & A1_ESTADO & "','" & A1_CEP & "','" & A1_DDI & "','" & A1_DDD _
                                     & "','" & A1_TEL & "','" & A1_TELEX & "','" & A1_FAX & "','" & A1_ENDENTNR _
                                     & "','" & A1_END & "','" & A1_CGC & "','" & A1_CGC & "','" & A1_CONTATO _
                                     & "','" & A1_INSCR & "','" & A1_SUFRAMA & "','" & A1_BAIRRO & "','" & A1_CEP _
                                     & "','" & A1_MUN & "','" & A1_EST & "','" & A1_BAIRRO & "','" & A1_CEPE _
                                     & "','" & A1_MUNE & "','" & A1_ESTE & "','" & A1_EMAIL & "'," & A1_CODCIDADE _
                                     & ",'" & A1_INSCRUR & "','" & A1_CODLOJA _
                                     & "'," & A1_CODCIDADE & "," & A1_CODCIDENT & ",'" & A1_UFENTREGA & "'"

            CONECTA_GLOBAL.Execute "EXEC " & SQL
      End If

      TabCliIntegra.MoveNext
   Wend
   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "CLIENTE_INTEGRA"
End Sub

Sub TRANSPORTADORA_INTEGRA(CNPJ_CPF_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(CNPJ_CPF_A) = "" Then _
      Exit Sub

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabCliIntegra    As New ADODB.Recordset
   Dim TabTempIntegra   As New ADODB.Recordset
   Dim MFTCOD           As Long
   Dim MFTNOME          As String
   Dim MFTNREDUZ        As String
   Dim MFTEND           As String
   Dim MFTBAIRRO        As String
   Dim MFTMUN           As String
   Dim MFTEST           As String
   Dim MFTCEP           As String
   Dim MFTDDD           As String
   Dim MFTTEL           As String
   Dim MFTCGC           As String
   Dim MFTINSEST        As String
   Dim MFTEMAIL         As String
   Dim MFTREGISTRO      As String
   Dim MFTCODCID        As String
   Dim MFTNATUREZA      As String
   Dim MFTRUA           As String

   A1_FILIAL = "0" & EMPRESA_ID_N

   If Len(CNPJ_CPF_A) <= 11 Then
      TIPO_PESSOA = "F"
      Else: TIPO_PESSOA = "J"
   End If

   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close

   SQL = "select TRANSPORTADORA.TRANSP_ID, TRANSPORTADORA.PESSOA_ID, TRANSPORTADORA.DT_CAD, TRANSPORTADORA.STATUS, "
   SQL = SQL & " TRANSPORTADORA.ESTABELECIMENTO_ID, PESSOA.CNPJCPF, PESSOA.DESCRICAO, PESSOA.RAZAO, PESSOA.DATA_CAD, "
   SQL = SQL & " FONE.NUMERO AS TeleFone, FONE.DDD, ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, ENDERECO.TIPO, "
   SQL = SQL & " ENDERECO.NUMERO AS NUMERO_ENDERECO, CEP.CEP_ID, CEP.CIDADE, CEP.UF, CEP.IBGE_ID"
   SQL = SQL & " from TRANSPORTADORA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   SQL = SQL & " ON TRANSPORTADORA.PESSOA_ID = PESSOA.PESSOA_ID "
   SQL = SQL & " INNER JOIN ENDERECO WITH (NOLOCK)"
   SQL = SQL & " ON PESSOA.PESSOA_ID = ENDERECO.PESSOA_ID "
   SQL = SQL & " INNER JOIN FONE WITH (NOLOCK)"
   SQL = SQL & " ON PESSOA.PESSOA_ID = FONE.PESSOA_ID "
   SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
   SQL = SQL & " ON ENDERECO.CEP_ID = CEP.CEP_ID"

   SQL = SQL & " where status = 'A' "
   SQL = SQL & " and tipo = 'C' "   'tipo do endereço que deve ser comercial

   If Trim(CNPJ_CPF_A) <> "" Then _
      SQL = SQL & " and cnpjcpf = '" & Trim(CNPJ_CPF_A) & "'"

   TabCliIntegra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabCliIntegra.EOF
      LIMPA_VARIAVEIS

A1_FILIAL = "0" & Trim(TabCliIntegra.Fields("estabelecimento_id").Value)

      MFTCOD = "" & Trim(TabCliIntegra.Fields("transp_id").Value)
      MFTNOME = "" & Left(Trim(TabCliIntegra.Fields("DESCRICAO").Value), 60)
      MFTNREDUZ = "" & Left(Trim(TabCliIntegra.Fields("DESCRICAO").Value), 60)
      MFTEND = "" & Left(Trim(TabCliIntegra.Fields("rua").Value), 60)
      MFTBAIRRO = "" & Left(Trim(TabCliIntegra.Fields("bairro").Value), 60)
      MFTMUN = "" & Left(Trim(TabCliIntegra.Fields("cidade").Value), 60)
      MFTEST = "" & Left(Trim(TabCliIntegra.Fields("uf").Value), 2)
      MFTCEP = "" & Left(Trim(TabCliIntegra.Fields("cep_id").Value), 8)
      MFTDDD = "" & Left(Trim(TabCliIntegra.Fields("ddd").Value), 3)
      MFTTEL = "" & Left(Trim(TabCliIntegra.Fields("telefone").Value), 15)
      MFTCGC = "" & Trim(CNPJ_CPF_A)
      MFTINSEST = "" & TRAZ_IE(TabCliIntegra.Fields("pessoa_id").Value)
      MFTEMAIL = "" & Left(Trim(TRAZ_EMAIL(TabCliIntegra.Fields("pessoa_id").Value)), 30)
      MFTREGISTRO = "" & Trim(TabCliIntegra.Fields("transp_id").Value)
      MFTCODCID = "" & Trim(TabCliIntegra.Fields("ibge_id").Value)
      MFTNATUREZA = "" & TIPO_PESSOA
      MFTRUA = "" & Left(Trim(TabCliIntegra.Fields("rua").Value), 12)

      If TabTempIntegra.State = 1 Then _
         TabTempIntegra.Close

      SQL = "select mftcgc,mftcod from MFT010 WITH (NOLOCK)"
      SQL = SQL & " where mftcgc = '" & Trim(CNPJ_CPF_A) & "'"
      TabTempIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabTempIntegra.EOF Then
         MFTCOD = 0 & TabTempIntegra.Fields("MFTCOD").Value
         Acao_N = 2
         Else
            Acao_N = 1
            MFTCOD = 1

            If TabConsulta.State = 1 Then _
               TabConsulta.Close
            SQL = "select max(mftregistro) from MFT010 WITH (NOLOCK)"
            TabConsulta.Open SQL, CONECTA_GLOBAL, , , adCmdText
            If Not TabConsulta.EOF Then _
                If Not IsNull(TabConsulta.Fields(0).Value) Then _
                    MFTCOD = TabConsulta.Fields(0).Value + 1
            If TabConsulta.State = 1 Then _
               TabConsulta.Close

      End If
      If TabTempIntegra.State = 1 Then _
         TabTempIntegra.Close

      SQL = "spTransportadoraGlobal " & Acao_N _
                               & ",'" & A1_FILIAL & "','" & MFTCOD & "','" & MFTNOME & "','" & MFTNREDUZ _
                               & "','" & MFTEND & "','" & MFTBAIRRO & "','" & MFTMUN & "','" & MFTEST _
                               & "','" & MFTCEP & "','" & MFTDDD & "','" & MFTTEL & "','" & MFTCGC _
                               & "','" & MFTINSEST & "','" & MFTEMAIL & "'," & MFTREGISTRO & ",'" & MFTCODCID _
                               & "','" & MFTNATUREZA & "','" & MFTRUA & "'"

      CONECTA_GLOBAL.Execute "EXEC " & SQL

      TabCliIntegra.MoveNext
   Wend
   If TabCliIntegra.State = 1 Then _
      TabCliIntegra.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "TRANSPORTADORA_INTEGRA"
End Sub

Function PEDIDO_INTEGRA_MFA010(NF_ID_N As Long, _
                               MFATRANSP As Long, _
                               MFAPREFIXO As String, _
                               MFAOBSNOTA As String, _
                               MFAINDFINAL As String, _
                               MFAIDDEST As String, _
                               MFAINDPRES As String, _
                               MFACHAVEREFNFE As String, _
                               MFAFINNFE As String, _
                               MFAPLIQUI As String, _
                               MFAPBRUTO As String, _
                               MFATIFRETE As String, _
                               MFACODSITT As String, _
                               MFAVALIMP5 As Double, _
                               MFAVALIMP6 As Double, _
                               MFAVOLUME1 As Double, _
                               MFATIPOREM As Integer, _
                               MFATIPO As String, _
                               MFAFRETE As Double) As Boolean
'On Error GoTo ERRO_TRATA

   PEDIDO_INTEGRA_MFA010 = False

   If NF_ID_N <= 0 Then
      MsgBox "Documento fiscal não encontrado !!!"
      Exit Function
   End If
   If Trim(MFAINDFINAL) = "" Then
      MsgBox "Indica operação com Consumidor final da NF-e não informado !!!"
      Exit Function
   End If
   If Trim(MFAIDDEST) = "" Then
      MsgBox "          da NF-e não informado !!!"
      Exit Function
   End If
   If Trim(MFAINDPRES) = "" Then
      MsgBox "Indicador de presença do comprador da NF-e não informado !!!"
      Exit Function
   End If
   If Trim(MFAFINNFE) = "" Then
      MsgBox "Finalidade de emissão da NF-e não informado !!!"
      Exit Function
   End If

   Dim TabCabeca           As New ADODB.Recordset
   Dim TabCabecaIntegra    As New ADODB.Recordset
   Dim MFASEQUENCIA_N      As Long
   Dim MFANOMECONSUMIDOR   As String
   Dim MFACPFCONSUMIDOR    As String
   Dim MFAFILIAL           As String
   Dim MFAVALBRUT          As Double
   Dim MFADOC              As String
   Dim MFACODEMP           As String
   Dim MFALOJA             As String
   Dim MFADTLANC           As String
   Dim MFADTBASE0          As String
   Dim MFADTBASE1          As String
   Dim MFADTENTR           As String
   Dim MFANFEEMISE         As String
   Dim MFADTENSAI          As String
   Dim MFADTDIGIT          As String
   Dim MFADTNSAI           As String
   Dim MFADTDIGT           As String
   Dim MFANFCUPOM          As String
   Dim MFACHAVENFE         As String
   Dim MFACODPROT          As String
   Dim MFANFECNF           As String
   Dim vFCPST              As Double
   Dim vFCPSTRet           As Double
   Dim MFAVALMERC          As String

'MFAICMFRET , MFAVALICM , MFABASEICM , MFAVALIPI , MFABASEIPI , MFABASEINS , MFABASICMST , MFAVALICMST

   Dim MFAICMFRET          As String   'ACHO QUE É ICMS SOBRE O FRETE
   Dim MFAVALICM           As String
   Dim MFABASEICM          As String
   Dim MFAVALIPI           As String
   Dim MFABASEIPI          As String
   Dim MFABASEINS          As String
   Dim MFABASICMST         As String
   Dim MFAVALICMST         As String
   
   MFABASEIPI = ""
   MFAVALIPI = ""
   MFAICMFRET = ""                     'ACHO QUE É ICMS SOBRE O FRETE
   MFABASEINS = ""
   MFABASICMST = ""
   MFAVALICMST = ""

   'ICMS
   MFABASEICM = ""
   MFAVALICM = ""

   Acao_N = 0
   MFADOC = ""
   PEDIDO_ID_N = 0

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   If TabCabeca.State = 1 Then _
      TabCabeca.Close

   SQL = "SELECT NF.*, PESSOA.* FROM NF WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN PESSOA WITH (NOLOCK) "
   SQL = SQL & " ON NF.PESSOA_ID = PESSOA.PESSOA_ID"

   SQL = SQL & " WHERE NF.nf_id = " & NF_ID_N

   TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabCabeca.EOF Then
      LIMPA_VARIAVEIS

      If Trim(TIPO_NFe_GERAR) = "R" Then
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close

         SQL = "select * from PEDIDONF"
         SQL = SQL & " where nf_id = " & TabCabeca.Fields("nf_id").Value
         TabCabecaIntegra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabCabecaIntegra.EOF Then
            PEDIDO_ID_N = 0 & TabCabecaIntegra.Fields("pedido_id").Value
            MFANFECNF = "" & PEDIDO_ID_N
         End If
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         Else: MFANFECNF = "" & TabCabeca.Fields("numr_nota").Value
      End If

      PESSOA_ID_N = 0 & TabCabeca.Fields("pessoa_id").Value
      
      MFADOC = "" & TabCabeca.Fields("numr_nota").Value
      MFASERIE = "1"
      MFACPFCONSUMIDOR = "" & Trim(TabCabeca.Fields("CNPJCPF").Value)
      MFACLIENTE = "" & TRAZ_ID_TABELA_GLOBAL("SA1010", "A1_COD", "A1_CGC", MFACPFCONSUMIDOR)

      MFALOJA = "0" & ESTABELECIMENTO_ID_N
      MFAFILIAL = "0" & ESTABELECIMENTO_ID_N

      MFANOMECONSUMIDOR = ""
      'MFACPFCONSUMIDOR = ""
      If Trim(MFACPFCONSUMIDOR) <> "99999999999" Then
         MFACPFCONSUMIDOR = "" & TabCabeca.Fields("CNPJCPF").Value
         MFANOMECONSUMIDOR = "" & Trim(Left(TabCabeca.Fields("descricao").Value, 60))
         Else
            'If TabCabecaIntegra.State = 1 Then
            '   TabCabecaIntegra.Close

            'SQL = "select nome_cliente from PEDIDO WITH (NOLOCK)"
            'SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
            'TabCabecaIntegra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            'If Not TabCabecaIntegra.EOF Then _
               MFANOMECONSUMIDOR = "" & Trim(Left(TabCabecaIntegra.Fields(0).Value, 60))
            MFACPFCONSUMIDOR = ""
      End If

'=============
      SQL = "delete MFA010 "
      SQL = SQL & " where mfadoc = '" & Trim(MFADOC) & "'"   'considera que a sequencia de nota fiscal é unica, por isso le dessa forma

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

      SQL = SQL & " and MFAPREFIXO = '" & Trim(MFAPREFIXO) & "'"
      CONECTA_GLOBAL.Execute SQL

      MFASEQUENCIA_N = 1
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close
      SQL = "select max(MFASEQUENCIA) from MFA010 WITH (NOLOCK)"
      TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabCabecaIntegra.EOF Then _
         If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
            MFASEQUENCIA_N = TabCabecaIntegra.Fields(0).Value + 1
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close

      MFACOND = "001"
      MFAEMISSAO = "" & TabCabeca.Fields("DT_EMISSAO").Value
      MFADTLANC = MFAEMISSAO
      MFADTBASE0 = MFAEMISSAO
      MFADTBASE1 = MFAEMISSAO
      MFADTENTR = MFAEMISSAO
      MFANFEEMISE = MFAEMISSAO
      MFADTENSAI = MFAEMISSAO
      MFADTDIGIT = MFAEMISSAO
      MFACODSITT = Trim(Left(MFACODSITT, 59))
      MFAEST = "52"
      'MFAFRETE = 0

      If Len(Trim(TabCabeca.Fields("CNPJCPF").Value)) = 11 Then
         MFATIPOCLI = "F"
         Else: MFATIPOCLI = "J"
      End If

      MFAVALBRUT = 0
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close
      SQL = "select sum((valor*qtde)-desconto) from NFITEM WITH (NOLOCK)"
      SQL = SQL & " where nf_id = " & NF_ID_N
      TabCabecaIntegra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCabecaIntegra.EOF Then _
         If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
            MFAVALBRUT = 0 & TabCabecaIntegra.Fields(0).Value
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close

      MFAVALMERC = "" & MFAVALBRUT
      MFAESPECI2 = "PROPRIO"

      If Trim(MFATRANSP) = "" Then _
         MFATRANSP = "1"

      MFAVALFAT = "" & MFAVALBRUT
      MFAESPECI1 = "UN"

      If Trim(MFAPREFIXO) = "" Then
         MFAPREFIXO = "NFE"
         Else
            If Trim(MFAPREFIXO) = "55" Then _
               MFAPREFIXO = "NFE"
            If Trim(MFAPREFIXO) = "65" Then _
               MFAPREFIXO = "NFC"
      End If
      MFAMOEDA = "1"

      MFAREGISTRO = "1"
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close
      SQL = "select max(MFAREGISTRO) from MFA010 WITH (NOLOCK)"
      TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabCabecaIntegra.EOF Then _
         If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
            MFAREGISTRO = TabCabecaIntegra.Fields(0).Value + 1
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close

      MFACODSTAT = "0"
      MFACODMORE = "0"
      MFAVALTOT = "" & MFAVALBRUT
      MFAVALLIQUI = "" & MFAVALBRUT

      MFANFCUPOM = ""
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close
      SQL = "select numr_cupom from CUPOM WITH (NOLOCK)"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabCabecaIntegra.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCabecaIntegra.EOF Then _
         If Not IsNull(TabCabecaIntegra.Fields(0).Value) Then _
            MFANFCUPOM = 0 & TabCabecaIntegra.Fields(0).Value
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close

      'CABEÇA DO PEDIDO
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close

      'MODELO_DOCUMENTO = "" & MFAPREFIXO  'NF-e modelo 55 ou NF-e modelo 65

      SQL = "select mfadoc,MFASEQUENCIA,MFAREGISTRO from MFA010 WITH (NOLOCK)"
      SQL = SQL & " where mfadoc = '" & Trim(MFADOC) & "'"   'considera que a sequencia de nota fiscal é unica, por isso le dessa forma

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

      SQL = SQL & " and MFAPREFIXO = '" & Trim(MFAPREFIXO) & "'"
      TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If TabCabecaIntegra.EOF Then
         Acao_N = 1
         Else  'update
            MFASEQUENCIA_N = 0 & TabCabecaIntegra.Fields("MFASEQUENCIA").Value
            MFAREGISTRO = "" & TabCabecaIntegra.Fields("MFAREGISTRO").Value
            Acao_N = 2
      End If
      If TabCabecaIntegra.State = 1 Then _
         TabCabecaIntegra.Close

         MFACHAVENFE = ""
         MFACODPROT = ""
         vFCPST = 0
         vFCPSTRet = 0
         MFACODEMP = "01"

         MFANFECNF = "" & MFASEQUENCIA_N

         SQL = Acao_N & _
               ",'" & MFACODEMP & "','" & MFADOC & "','" & MFASERIE & "','" & MFACLIENTE & "'" & _
               ",'" & MFALOJA & "','" & MFASEQUENCIA_N & "','" & MFACOND & "'" & _
               ",'" & MFAEMISSAO & "','" & MFAEST & "','" & Replace(MFAFRETE, ",", ".") & "'" & _
               ",'" & MFATIPOCLI & "','" & Replace(MFAVALBRUT, ",", ".") & "','" & Replace(MFAVALMERC, ",", ".") & "'" & _
               ",'" & MFATIPO & "','" & MFAESPECI1 & "','" & MFAESPECI2 & "'" & _
               ",'" & Replace(MFAVOLUME1, ",", ".") & "','" & Replace(MFAPLIQUI, ",", ".") & "','" & Replace(MFAPBRUTO, ",", ".") & "'" & _
               ",'" & MFATRANSP & "','" & MFADTLANC & "','" & MFADTBASE0 & "'" & _
               ",'" & MFADTBASE1 & "','" & MFAFILIAL & "','" & Replace(MFAVALFAT, ",", ".") & "'" & _
               ",'" & MFAPREFIXO & "','" & MFAMOEDA & "','" & MFADTENTR & "'" & _
               ",'" & MFAREGISTRO & "','" & MFANFEEMISE & "'" & _
               ",'" & MFACODSTAT & "','" & MFACODMORE & "','" & MFADTENSAI & "'" & _
               ",'" & Replace(MFATIFRETE, ",", ".") & "','" & MFADTDIGIT & "','" & Replace(MFAVALTOT, ",", ".") & "'" & _
               ",'" & Replace(MFAVALLIQUI, ",", ".") & "','" & MFAINDFINAL & "'" & _
               ",'" & MFAIDDEST & "','" & MFAINDPRES & "','" & MFACHAVEREFNFE & "'" & _
               ",'" & MFAFINNFE & "','" & Replace(MFANOMECONSUMIDOR, ",", ".") & "','" & Replace(MFACPFCONSUMIDOR, ",", ".") & "'" & _
               ",'" & Replace(MFACODSITT, ",", ".") & "','" & MFANFCUPOM & "','" & Replace(MFAOBSNOTA, ",", ".") & "'" & _
               ",'" & Replace(MFAVALIMP5, ",", ".") & "','" & Replace(MFAVALIMP6, ",", ".") & "','" & MFATIPOREM & "'" & _
               ",'" & MFACHAVENFE & "','" & MFACODPROT & "','" & MFANFECNF & "'" & _
               ",'" & Replace(vFCPST, ",", ".") & "','" & Replace(vFCPSTRet, ",", ".") & "'"

      CONECTA_GLOBAL.Execute "EXEC spMFA010Global " & SQL
   End If
   If TabCabeca.State = 1 Then _
      TabCabeca.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   PEDIDO_INTEGRA_MFA010 = True

   If PEDIDOitem_INTEGRA_MFI010(MFASEQUENCIA_N, MFADOC, MFAPREFIXO, MFACLIENTE, MFAEMISSAO) = True Then
      frmINTEGRA.FINANCEIRO_INTEGRA MFAPREFIXO, MFADOC

      SQL = "update NF set "
         SQL = SQL & " status = 'E' "
      SQL = SQL & " where nf_id = " & NF_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      Else
         MsgBox "Não passou MFI010 " & PEDIDO_ID_N
         TRATA_ERROS "Não gravou item pedido = " & PEDIDO_ID_N, Me.Name, "PEDIDOitem_INTEGRA_MFI010"
   End If

   '====ROTINA REGIME NORMAL
   '====ROTINA REGIME NORMAL
   '====ROTINA REGIME NORMAL
   If CTR_EMPRESA_N = 3 Then
   'pegando totais icms tabela PEDIDOITEM
      If TabCabeca.State = 1 Then _
         TabCabeca.Close

      SQL = "select sum(valor_item) as BaseIcms, sum(vlricms) as ValorIcms from PEDIDOITEM"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabCabeca.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCabeca.EOF Then
         MFABASEICM = 0 & TabCabeca.Fields("BaseIcms").Value
         MFAVALICM = 0 & TabCabeca.Fields("ValorIcms").Value
      End If
      If TabCabeca.State = 1 Then _
         TabCabeca.Close

   'atualizando campos para o ICMS
      If CONECTA_GLOBAL.State <> 1 Then _
         ABRE_BANCO_GLOBAL

      SQL = "update MFA010 set "

      SQL = SQL & " MFABASEICM = '" & tpMOEDA(MFABASEICM) & "'"
      SQL = SQL & ",MFAVALICM = '" & tpMOEDA(MFAVALICM) & "'"

      SQL = SQL & " where mfadoc = '" & Trim(MFADOC) & "'"   'considera que a sequencia de nota fiscal é unica, por isso le dessa forma

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

      SQL = SQL & " and MFAPREFIXO = '" & Trim(MFAPREFIXO) & "'"
      CONECTA_GLOBAL.Execute SQL

      If CONECTA_GLOBAL.State = 1 Then _
         CONECTA_GLOBAL.Close
   End If
   '====================================================

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PEDIDO_INTEGRA_MFA010"
End Function

Function PEDIDOitem_INTEGRA_MFI010(MFISEQUEN_N As Long, _
                               MFIDOC_A As String, _
                               MFAPREFIXO_MFI010_A As String, _
                               MFICLIENTE As String, _
                               MFIEMISSAO As String) As Boolean
'On Error GoTo ERRO_TRATA

   PEDIDOitem_INTEGRA_MFI010 = False

   If Trim(MFIDOC_A) = "" Then
      MsgBox "Documento fiscal não encontrado !!!"
      Exit Function
   End If
   If Trim(MFAPREFIXO_MFI010_A) = "" Then
      MsgBox "PreFixo(NFe ou NFCe) não informado !!!"
      Exit Function
   End If

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabItem       As New ADODB.Recordset
   Dim TabIntegra    As New ADODB.Recordset

   Dim MFIFILIAL     As String
   Dim MFILOJA_A     As String
   Dim MFIITEM       As Long
   Dim MFICOD        As String
   Dim MFIUM         As String
   Dim MFIQUANT      As Double
   Dim MFIPRCVEN     As Double
   Dim MFITOTAL      As Double
   Dim MFIVALIPI     As Double
   Dim MFIVALICM     As Double
   Dim MFITES        As String
   Dim MFICF         As String
   Dim MFIDESC       As Double
   Dim MFIIPI        As Double
   Dim MFIPICM       As Double
   Dim MFIPESO       As Double
   Dim MFILOCAL_A    As String
   Dim MFIGRUPO      As String
   Dim MFITP         As String
   Dim MFISERIE_A    As String
   Dim MFIEST        As String
   Dim MFIDESCON     As Double
   Dim MFITIPO       As String
   Dim MFIQTDEDEV    As Double
   Dim MFIVALDEV     As Double
   Dim MFIORIGLAN    As String
   Dim MFIDTLCTCT    As String
   Dim MFICLASFIS    As String
   Dim MFIQTDEFAT    As Double
   Dim MFIQTDAFAT    As Double
   Dim MFISITRIB     As String
   Dim MFIPESLIQ     As Double
   Dim MFIPESBRU     As Double
   Dim MFIVALLIQ     As Double
   Dim MFINFORI      As String
   Dim MFIDESTOTIT   As Double
   Dim MFIALIICMS    As Double
   Dim MFIBASICMST   As Double
   Dim MFIALIICMST   As Double
   Dim MFIVALICMST   As Double
   Dim MFIVALBRUT    As Double
   Dim MFIVALBONI    As Double
   Dim MFIVALTROCA   As Double
   Dim MFIQTDVOL     As Double
   Dim CFOP_ID_N     As Integer

Dim STRIBUTARIA_A As String

   Dim vBCFCPSTRet   As Double
   Dim pFCPSTRet     As Double
   Dim vFCPSTRet     As Double
   Dim MFICEAN       As String
   Dim MFICEANTRIB   As String
   Dim MFIBASEICM    As Double   'BASE DE CALCULO ICMS
   Dim BASE_CALCULO_REDUZIDA  As Double
   Dim PERC_BASE_REDUZIDA     As Double
   Dim MFIALIICMRED  As Double

   Dim PISCST_N      As Double
   Dim COFINSCST_N   As Double
   Dim PISVBC_N            As Double   '1 - PISVBC VALOR BSASE DE CALUILO DO PIS
   Dim PISVPIS_N           As Double   '3 - PISVPIS = VALOR DO PIS APLICAR OM PERCENTAUL SOBRE A BASE PEDE A CONTADORA PARA SABER SABER A BASE DE CACULO
'---PARA COFINS :   MESMO CRITERIO DE CONCEITOS
   Dim COFINSVBC_N         As Double   '1 - COFINSVBC
   Dim COFINSVCOFINS_N     As Double   '3 - COFINSVCOFINS
'---CRIAR MAIS ESTES CAMPOS NO MFI010 :
   Dim PisqBCProd_A        As String   'PisqBCProd
   Dim PisvAliqProd_A      As String   'PisvAliqProd
   Dim COFINSqBCProd_A     As String   'COFINSqBCProd
   Dim COFINSvAliqProd_A   As String   'COFINSvAliqProd

   If Trim(UF_EMPRESA_A) = "" Then _
      PEGA_DADOS_EMPRESA
   'aqui é ajustado se for consumidor final tem que pegar o mesmo UF de destrino para aliquotas
   'If Trim(txtCNPJCPF.Text) = "99999999999" Then _
      UF_CLIENTE_A = "" & UF_EMPRESA_A

   If Trim(UF_CLIENTE_A) = "" Then _
      UF_CLIENTE_A = "" & UF_EMPRESA_A

   BASE_CALCULO_REDUZIDA = 0
   PERC_BASE_REDUZIDA = 0

   PISVBC_N = 0
   ALIQUTOA_PIS_N = 0
   PISVPIS_N = 0

   COFINSVBC_N = 0
   ALIQUTOA_COFINS_N = 0
   COFINSVCOFINS_N = 0

   PisqBCProd_A = ""
   PisvAliqProd_A = ""
   COFINSqBCProd_A = ""
   COFINSvAliqProd_A = ""

   SQL = "delete MFI010 "
   SQL = SQL & " where MFISEQUEN = " & MFISEQUEN_N
   CONECTA_GLOBAL.Execute SQL

   MFIFILIAL = "01"
   MFILOCAL_A = "0" & ESTABELECIMENTO_ID_N
   MFILOJA_A = "0" & ESTABELECIMENTO_ID_N
   MFISERIE_A = "" & MFAPREFIXO_MFI010_A

   SQL = "delete MFI010 "
   SQL = SQL & " where mfidoc = '" & Trim(MFIDOC_A) & "'"
   SQL = SQL & " and mfiloja = '" & Trim(MFILOJA_A) & "'"
   SQL = SQL & " and mfilocal = '" & Trim(MFILOCAL_A) & "'"
   SQL = SQL & " and mfiserie = '" & Trim(MFISERIE_A) & "'"
   CONECTA_GLOBAL.Execute SQL

   Acao_N = 0

   If TabItem.State = 1 Then _
      TabItem.Close

   SQL = "SELECT NF.PESSOA_ID, NF.TRANSP_ID, NF.NF_TIPO, "
   SQL = SQL & " NF.NUMR_NOTA, NF.SERIE_NOTA, NF.DT_EMISSAO, NF.STATUS, "
   SQL = SQL & " NF.ESTABELECIMENTO_ID, NF.MODELO_DOC, NFITEM.*,"
   SQL = SQL & " UNIDADE_MEDIDA, SITUACAO_TRIBUTARIA, ALIQUOTA_ICMS, CODG_PRODUTO "

   SQL = SQL & " FROM NF WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN NFITEM WITH (NOLOCK)"
   SQL = SQL & " ON NF.NF_ID = NFITEM.NF_ID"
   SQL = SQL & " INNER JOIN PRODUTO "
   SQL = SQL & " ON NFITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"

   SQL = SQL & " where numr_nota = " & MFIDOC_A
   SQL = SQL & " and modelo_doc = '" & Trim(MFAPREFIXO_MFI010_A) & "'"
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabItem.EOF
      'MFIFILIAL = "0" & EMPRESA_ID_N
      MFIFILIAL = "01"
      MFILOCAL_A = "0" & ESTABELECIMENTO_ID_N
      MFILOJA_A = "0" & ESTABELECIMENTO_ID_N

      MFIITEM = 1
      If TabIntegra.State = 1 Then _
         TabIntegra.Close
      SQL = "select max(MFIITEM) from MFI010 WITH (NOLOCK)"
      SQL = SQL & " where MFISEQUEN  = " & MFISEQUEN_N     'vem do MFA010.MFASEQUENCIA   É O CAMPO DE RELAÇÃO ENTRE AS DUAS TABELAS
      TabIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabIntegra.EOF Then
         If Not IsNull(TabIntegra.Fields(0).Value) Then
            MFIITEM = 0 & TabIntegra.Fields(0).Value
            MFIITEM = MFIITEM + 1
         End If
      End If
      If TabIntegra.State = 1 Then _
         TabIntegra.Close

      'lendo tabela de produtos
      MFICOD = ""
      If TabIntegra.State = 1 Then _
         TabIntegra.Close
      SQL = "select b1_cod from SB1010 WITH (NOLOCK)"
      SQL = SQL & " where b1_codant = '" & Trim(TabItem.Fields("codg_produto").Value) & "'"
      TabIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabIntegra.EOF Then _
         If Not IsNull(TabIntegra.Fields(0).Value) Then _
            MFICOD = "" & Trim(TabIntegra.Fields(0).Value)
      If TabIntegra.State = 1 Then _
         TabIntegra.Close

      MFIUM = "" & Trim(TabItem.Fields("unidade_medida").Value)
      MFIQUANT = 0 & Trim(TabItem.Fields("qtde").Value)
      MFIPRCVEN = 0 & Trim(TabItem.Fields("valor").Value)
      MFITOTAL = 0 & (TabItem.Fields("valor").Value * TabItem.Fields("qtde").Value)
      MFIVALIPI = 0
      MFITES = "001"
      MFICF = "" & TabItem.Fields("cfop_id").Value
      MFIDESC = 0
      MFIIPI = 0
      MFIPICM = 0
      MFIPESO = 0
      MFIGRUPO = "0001"
      MFITP = "MP"
      MFISERIE_A = "" & MFAPREFIXO_MFI010_A
      MFIEST = "52"
      MFIDESCON = 0
      MFITIPO = "N"
      MFIQTDEDEV = 0 & Trim(TabItem.Fields("qtde").Value)
      MFIVALDEV = 0 & (TabItem.Fields("valor").Value * TabItem.Fields("qtde").Value)
      MFIORIGLAN = "NF" '& MFAPREFIXO
      MFIDTLCTCT = "" & Trim(TabItem.Fields("DT_EMISSAO").Value)

      '========================
      '========================
      '========================
      'Segue uma pequena orientação sobre o uso do CFOP e do CSOSN.
      '- Quando efetuar a revenda dentro do Estado sem substituição tributária usar.
      'CFOP: 5.102 podendo usar o CSOSN 0101 com permissão de crédito ou 0102 sem permissão de crédito quando efetuar a venda para pessoa física ou pessoa jurídica.
      '- Quando efetuar uma revenda para fora do estado sem substituição tributária usar.
      'CFOP: 6102 - podendo usar o CSOSN 0101 com permissão de crédito ou 0102 sem permissão de crédito quando efetuar a venda para pessoa física ou pessoa jurídica.
      '- Quando efetuar uma revenda para dentro do estado com substituição tributária usar.
      'CFOP: 5405 e o CSOSN 0500.
      'Quando efetuar uma revenda para fora do estado com substituição tributária usar.
      'CFOP: 6404 e o CSOSN 0500.

      '12 Tributação do ICMS pelo Simples Nacional sem permissao  102
      '15 Tributação do ICMS pelo Simples Nacional(500)           500
      '16 Tributação do ICMS pelo Simples Nacional(900)           900
      '17 Tributação do ICMS pelo Simples Nacional Nao tributado  400

'------------------------------------------------------------------
'CST         'Tributação do ICMS                      'MFICLASFIS

'vem dessa relação ai:
   'INNER JOIN MFTCLASFISTRI
   'ON MFI010.MFICLASFIS = MFTCLASFISTRI.MFTCODIGO

   'onde:
      'select * from MFTCLASFISTRI =>
'0  TRIBUTADA INTEGRALMENTE                                                         000
'1  TRIBUTADA E COM COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA                    010
'2  COM REDUCAO DE BASE DE CALCULO                                                  020
'3  ISENTA OU NAO TRIBUTADA E COM COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA      030
'4  ISENTA                                                                          040
'5  NAO TRIBUTADA                                                                   041
'6  SUSPENSAO                                                                       050
'7  DIFERIMENTO                                                                     051
'8  ICMS COBRADO ANTERIORMENTE POR SUBSTITUICAO TRIBUTARIA                          060
'9  COM REDUCAO DE BASE DE CALCULO E COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA   070
'10 OUTRAS                                                                          090

''''''''''''''''''''SIMPLES NACIONAL DAQUI PRA BAIXO
'11 Tributação do ICMS pelo Simples Nacional com Permissão de Crédito               101
   '12 Tributação do ICMS pelo Simples Nacional sem permissao                          102
'13 Tributação do ICMS pelo Simples Nacional(201)                                   201
'14 Tributação do ICMS pelo Simples Nacional(202 )                                  202
   '15 Tributação do ICMS pelo Simples Nacional(500)                                   500
'16 Tributação do ICMS pelo Simples Nacional(900)                                   900
   '17 Tributação do ICMS pelo Simples Nacional Nao tributado                          400
'18 Tributação do ICMS pelo Simples Nacional Isencao do ICMS                        103
'19 Tributação do ICMS pelo Simples Nacional Imune                                  300
'20 Tributação do ICMS pelo Simples Nacional(203)                                   203
'------------------------------------------------------------------
'INNER JOIN MTSITTRIBU
'ON MTSITTRIBU.MTSCODIGO = MFI010.MFISITRIB
   '8  Dev Venda Merc Adq ou Receb Terc.     1202  CFOP
   '42 VENDA DENTRO ESTADO                   5102  CFOP
   '57 VENDA FORA ESTADO                     6102  CFOP

'------------------------------------------------------------------
'INNER JOIN MTSITTRIBU
'ON MTSCODIGO = MFI010.MFISITRIB

'MTSCODFIS DQUI PEGA O CFOP DO ITEM NO GLOBAL
      MFISITRIB = "42"  'relacionado ao CFOP
      '42 VENDA DENTRO ESTADO  5102        S           1202  1

      If TabIntegra.State = 1 Then _
         TabIntegra.Close
      SQL = "select MTSCODIGO from MTSITTRIBU WITH (NOLOCK)"
      SQL = SQL & " where MTSCODFIS = '" & Trim(MFICF) & "'"
      TabIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabIntegra.EOF Then _
         If Not IsNull(TabIntegra.Fields(0).Value) Then _
            MFISITRIB = TabIntegra.Fields(0).Value
      If TabIntegra.State = 1 Then _
         TabIntegra.Close
'===================================================

'orig        'Origem da mercadoria
'modBC       'Modalidade de determinação da BC do ICMS
'vBC         'Valor da BC do ICMS
'pICMS       'Alíquota do imposto
'vICMS       'Valor do ICMS

      'ST_PRODUTO = "" & Trim(TabItem.Fields("SITUACAO_TRIBUTARIA").Value)
      'SE O ITEM FOR SUBSTITUIÇÃO TRIBUTARIA PASSA 500
      'If ST_PRODUTO = "10" Or ST_PRODUTO = "60" Then _
         MFICLASFIS = "15" 'select * from MFTCLASFISTRI where MFTCODIGO = 12 or MFTCODIGO = 17

      'If INSCRICAO_UF_A = "" Or INSCRICAO_UF_A = "ISENTO" Then
      '   If MFAPREFIXO = "NFC" Then
            'MFICLASFIS = "12"   'select * from MFTCLASFISTRI where MFTCODIGO = 12 or MFTCODIGO = 17
      '      Else: MFICLASFIS = "17" 'NFE 'select * from MFTCLASFISTRI where MFTCODIGO = 12 or MFTCODIGO = 17
      '   End If
      'End If
      '========================
      '========================
      '========================
      MFIQTDAFAT = 0 & Trim(TabItem.Fields("qtde").Value)
      MFIPESLIQ = 0
      MFIPESBRU = 0
      MFIVALLIQ = 0 & (TabItem.Fields("valor").Value * TabItem.Fields("qtde").Value)

      MFIDESTOTIT = "0"
      MFIBASICMST = "0"
      MFIALIICMST = "0"
      MFIVALICMST = "0"
      MFIALIICMRED = "0"
      MFIVALBRUT = "0"
      MFIVALBONI = "0"
      MFIVALTROCA = "0"
      MFIQTDVOL = "0"

      vBCFCPSTRet = 0
      pFCPSTRet = 0
      vFCPSTRet = 0
      MFICEAN = ""
      MFICEANTRIB = ""

      If Trim(MOSTRA_VERSAO_NFe(CNPJ_EMPRESA_N)) = "40" Then
         MFICEAN = "SEM GTIN"
         MFICEANTRIB = "SEM GTIN"
      End If

      MFIBASEICM = 0
      MFIALIICMS = 0
      MFIVALICM = 0

      MFICLASFIS = "17" 'NFE  'select * from MFTCLASFISTRI where MFTCODIGO = 12 or MFTCODIGO = 17
      '12 Tributação do ICMS pelo Simples Nacional sem permissao   102
      '17 Tributação do ICMS pelo Simples Nacional Nao tributado   400

      If Trim(MFAPREFIXO_MFI010_A) = "NFC" Then 'SIMPLES NACIONAL
         MFICLASFIS = "12" 'select * from MFTCLASFISTRI where MFTCODIGO = 12 or MFTCODIGO = 17
         MFICF = "5102"
      End If
'===================================================

      STRIBUTARIA_A = "" & TabItem.Fields("STRIBUTARIA").Value

'SIMPLES NACIONAL SIMPLES NACIONAL SIMPLES NACIONAL SIMPLES NACIONAL SIMPLES NACIONAL
      'Select Case STRIBUTARIA_A  'ESSA VARIAVEL VEM DO PRODUTO COM A SITUAÇÃO TRIBUTARIA
      '   Case "00"   '0-TRIBUTADA INTEGRALMENTE                                                        000
      '      MFICLASFIS = "101"
      '   Case "10"   '1  TRIBUTADA E COM COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA                  010
      '      MFICLASFIS = "1"
      '   Case "20"   '2  COM REDUCAO DE BASE DE CALCULO                                                020
      '      MFICLASFIS = "2"
      '   Case "30"   '3  ISENTA OU NAO TRIBUTADA E COM COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA    030
      '      MFICLASFIS = "3"
      '   Case "40"   '4  ISENTA                                                                        040
      '      MFICLASFIS = "4"
      '   Case "41"   '5  NAO TRIBUTADA                                                                 041
      '      MFICLASFIS = "5"
      '   Case "50"   '6  SUSPENSAO                                                                     050
      '      MFICLASFIS = "6"
      '   Case "51"   '7  DIFERIMENTO                                                                   051
      '      MFICLASFIS = "7"
      '   Case "60"   '8  ICMS COBRADO ANTERIORMENTE POR SUBSTITUICAO TRIBUTARIA                        060
      '      MFICLASFIS = "8"
      '   Case "70"   '9  COM REDUCAO DE BASE DE CALCULO E COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA 070
      '      MFICLASFIS = "9"
      '   Case "90"   '10 OUTRAS                                                                        090
      '      MFICLASFIS = "10"
      'End Select
'FIM SIMPLES NACIONAL SIMPLES NACIONAL SIMPLES NACIONAL SIMPLES NACIONAL

      '====ROTINA REGIME NORMAL
      '====ROTINA REGIME NORMAL
      '====ROTINA REGIME NORMAL
      MFINFORI = ""
      CST_ORIG_ICMS_N = 0

      If CTR_EMPRESA_N = 3 Then
         'MFIBASEICM = 0 & (TabItem.Fields("valor").Value * TabItem.Fields("qtde").Value)
         MFIBASEICM = 0 & MFITOTAL
         MFIALIICMS = 0 & TabItem.Fields("percicms").Value  'veio do prapara_item_tributaçao
         'MFIVALICM = 0 & TabItem.Fields("vlricms").Value
         MFIVALICM = 0 & MFITOTAL * MFIALIICMS / 100

         PISVBC_N = MFITOTAL  'PISVBC VALOR BSASE DE CALUILO DO PIS
         COFINSVBC_N = MFITOTAL

            'Call BUSCA_ALIQUOTA_PISCOFINS(UF_EMPRESA_A, UF_CLIENTE_A, MFICF)
            CFOP_ID_N = 0 & MFICF
            Call BUSCA_ALIQUOTA_PISCOFINS(UF_EMPRESA_A, "", CFOP_ID_N)

         PISVPIS_N = PISVBC_N * ALIQUTOA_PIS_N / 100
         COFINSVCOFINS_N = COFINSVBC_N * ALIQUTOA_COFINS_N / 100
         
         '1 - MFINFORI = Origem --> os dados antigos podem continuar com vazio(fiiz tratamento);
         MFINFORI = "" & CST_ORIG_ICMS_N

'ICMS
         CST_ICMS_A = Trim(CST_ICMS_A)
         Select Case CST_ICMS_A
            Case "00"   '0-TRIBUTADA INTEGRALMENTE                                                        000
               MFICLASFIS = "0"
            Case "10"   '1  TRIBUTADA E COM COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA                  010
               MFICLASFIS = "1"
               'Supondo, por exemplo, uma mercadoria com valor de R$ 1,00, com origem no estado do Rio de Janeiro,
               'e que vá ser vendida em São Paulo. Se sob essa operação incidir substituição tributária na cobrança do ICMS,
               'o governo estipulará uma pauta (isto é, um valor presumido de revenda - por exemplo, R$ 2,00).
               'Supondo que sob a operação interestadual entre SP e RJ incida uma alíquota de ICMS de 12%,
               'e a alíquota interna seja de 18%, o total de ICMS será calculado da seguinte maneira:

               'Total-ICMS = Valor-de-venda*ICMS interestadual + Valor-da-pauta*ICMS interno;
               'No nosso exemplo, os números seriam os seguintes:

               'Total-ICMS Normal = (R$1,00 * 12%) = 0,12
               'Total-ICMS Substituição =(R$2,00 * 18%) = 0,36
               'Como o ICMS é calculado como um debito e credito, ficaria assim o valor recolhido:

               '0,36 - (0,12) = R$ 0,24
               'O ICMS substituído se deduz do ICMS pago normalmente. Esse valor seria lançado na Nota Fiscal, e cobrado do cliente por duplicata.

               'Caso o emissor da Nota Fiscal não pague o ICMS (R$ 0,12) no prazo, ela será tachado de inadimplente.
               'Caso ela não pague o ICMS substitutivo (R$ 0,24) no prazo, além de inadimplente,
               'ele será processado como depositário infiel, estando seus responsáveis sujeitos até à prisão
               '(hoje em dia não estão mais sujeitos a prisão devido o
               'Pacto de San Jose da Costa Rica - adotado pelo Brasil e que manteve como única forma de prisão civil
               'para o devedor de pensão alimentícia). A responsabilidade do emissor independe da solvência do seu
               'cliente, ou seja, ele será considerado depositário infiel ainda que seu cliente não tenha pago a
               'nota emitida.
            Case "20"   '2  COM REDUCAO DE BASE DE CALCULO                                                020
               MFICLASFIS = "2"
               '2 - MFIALIICMRED = percentual de reducao na base(colocar a liquota para este CST ICMS: 020),
               'o resto dos campos e igual ao CST 00
               MFIALIICMRED = "" & PERC_BASE_REDUZ_N
               BASE_CALCULO_REDUZIDA = (MFIBASEICM * PERC_BASE_REDUZ_N) / 100

               'Base de Cálculo R$ 1.000,00
               'Redução da Base de Cálculo: 50%
               'Base de Cálculo Reduzida = R$ 1.000,00 - (50% de R$ 1.000,00) = R$ 1.000,00 - R$ 500,00 = R$ 500.00.
               'cst = 20
               'pRedBC = campo MFIALIICMRED   AS pRedBC
               'vBC = o memso do cst 00
               'vICMS = o mesmo do cst 00

               'cst = 30 e 60
               'modBCST = vem da tabela mftclassitri =
               'sb.AppendLine("(case when right(SituacaoTrib.MFTCODFIS,2)='10'
               'or right(SituacaoTrib.MFTCODFIS,2)='30'
               'OR right(SituacaoTrib.MFTCODFIS,2)='70'
               'OR right(SituacaoTrib.MFTCODFIS,2)='90' then '5' else   'N' end) as modBCST,");

               'pICMSST
               'vICMSST = MFIVALICMST
               'vBCST = MFIBASICMST
            Case "30"   '3  ISENTA OU NAO TRIBUTADA E COM COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA    030
               MFICLASFIS = "3"
            Case "40"   '4  ISENTA                                                                        040
               MFICLASFIS = "4"
            Case "41"   '5  NAO TRIBUTADA                                                                 041
               MFICLASFIS = "5"
            Case "50"   '6  SUSPENSAO                                                                     050
               MFICLASFIS = "6"
            Case "51"   '7  DIFERIMENTO                                                                   051
               MFICLASFIS = "7"
            Case "60"   '8  ICMS COBRADO ANTERIORMENTE POR SUBSTITUICAO TRIBUTARIA                        060
               MFICLASFIS = "8"
'OBS2 - CARA TA FACIL FECHAR A PARTE DO IMPOSTO O UNICO ESTADO QUE  VAI TER SUBSTITUIÇÃO TRINUTARIA E
'FORTALEZA-ce CST =60 SÃO DOIS CAMPOS A SABER :
'TABELA MFI010

'1 - MFIBASICMST : BASE DE CALCUO DA SUBSTITUIÇÃO TRIBUTARIA
'2 - MFIVALICMST VALOR DO ICMS SUBSTITUIÇÃO TRIBUTARIA

'O GRUPO CST = 060 SÓ PRECISA DESSES DOIS VALORES PREENCHIDOS

'E SO O QUE VOCE PRECISA FAZER PARA INTYEGRAR TUDO RELACIONADO AO ICMS

            Case "70"   '9  COM REDUCAO DE BASE DE CALCULO E COBRANCA DO ICMS POR SUBSTITUICAO TRIBUTARIA 070
               MFICLASFIS = "9"
            Case "90"   '10 OUTRAS                                                                        090
               MFICLASFIS = "10"
            Case "300"
               MFICLASFIS = "87"
         End Select
      End If

'PIS
         CST_PIS_A = Trim(CST_PIS_A)
         PIS_MEGASIM_GLOBAL
         
'COFINS
         CST_COFINS_A = Trim(CST_COFINS_A)
         COFINS_MEGASIM_GLOBAL

      If TabIntegra.State = 1 Then _
         TabIntegra.Close

      SQL = "select MFiSEQUEN from MFi010 WITH (NOLOCK)"
      SQL = SQL & " where MFISEQUEN = " & MFISEQUEN_N  'vem do MFA010.MFASEQUENCIA   É O CAMPO DE RELAÇÃO ENTRE AS DUAS TABELAS
      SQL = SQL & " and MFIITEM = " & MFIITEM
      TabIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If TabIntegra.EOF Then
         Acao_N = 1
         Else  'update
            MFISEQUEN_N = "" & TabIntegra.Fields("MFISEQUEN").Value   'vem do MFA010.MFASEQUENCIA   É O CAMPO DE RELAÇÃO ENTRE AS DUAS TABELAS
            Acao_N = 2
      End If
      If TabIntegra.State = 1 Then _
         TabIntegra.Close

         SQL = Acao_N & _
               ",'" & MFIFILIAL & "','" & MFIITEM & "','" & MFICOD & "'" & _
               ",'" & Replace(MFIUM, ",", ".") & "','" & Replace(MFIQUANT, ",", ".") & "','" & Replace(MFIPRCVEN, ",", ".") & "'" & _
               ",'" & Replace(MFITOTAL, ",", ".") & "','" & Replace(MFIVALIPI, ",", ".") & "','" & Replace(MFIVALICM, ",", ".") & "'" & _
               ",'" & Replace(MFITES, ",", ".") & "','" & Replace(MFICF, ",", ".") & "','" & Replace(MFIDESC, ",", ".") & "'" & _
               ",'" & Replace(MFIIPI, ",", ".") & "','" & Replace(MFIPICM, ",", ".") & "','" & Replace(MFIPESO, ",", ".") & "'" & _
               ",'" & MFICLIENTE & "','" & MFILOJA_A & "'" & _
               ",'" & MFILOCAL_A & "','" & MFIDOC_A & "','" & MFIEMISSAO & "'" & _
               ",'" & MFIGRUPO & "','" & MFITP & "','" & MFISERIE_A & "'" & _
               ",'" & MFIEST & "','" & Replace(MFIDESCON, ",", ".") & "'" & _
               ",'" & MFITIPO & "','" & Replace(MFIQTDEDEV, ",", ".") & "','" & Replace(MFIVALDEV, ",", ".") & "'" & _
               ",'" & MFIORIGLAN & "','" & MFIDTLCTCT & "','" & MFICLASFIS & "'" & _
               ",'" & Replace(MFIQTDEFAT, ",", ".") & "','" & Replace(MFIQTDAFAT, ",", ".") & "','" & MFISEQUEN_N & "'" & _
               ",'" & MFINFORI & "','" & MFISITRIB & "'" & _
               ",'" & Replace(MFIPESLIQ, ",", ".") & "','" & Replace(MFIPESBRU, ",", ".") & "','" & Replace(MFIVALLIQ, ",", ".") & "'" & _
               ",'" & Replace(MFIDESTOTIT, ",", ".") & "','" & Replace(MFIALIICMS, ",", ".") & "','" & Replace(MFIBASICMST, ",", ".") & "'" & _
               ",'" & Replace(MFIALIICMST, ",", ".") & "','" & Replace(MFIVALICMST, ",", ".") & "','" & Replace(MFIALIICMRED, ",", ".") & "'" & _
               ",'" & Replace(MFIVALBRUT, ",", ".") & "','" & Replace(MFIVALBONI, ",", ".") & "','" & Replace(MFIVALTROCA, ",", ".") & "'" & _
               ",'" & Replace(MFIQTDVOL, ",", ".") & "','" & Replace(vBCFCPSTRet, ",", ".") & "','" & Replace(pFCPSTRet, ",", ".") & "'" & _
               ",'" & Replace(vFCPSTRet, ",", ".") & "','" & Replace(MFICEAN, ",", ".") & "','" & Replace(MFICEANTRIB, ",", ".") & "'" & _
               ",'" & Replace(MFIBASEICM, ",", ".") & "','" & CST_PIS_A & "','" & CST_COFINS_A & "'" & _
               ",'" & Replace(PISVBC_N, ",", ".") & "','" & Replace(ALIQUTOA_PIS_N, ",", ".") & "','" & Replace(PISVPIS_N, ",", ".") & "'" & _
               ",'" & Replace(COFINSVBC_N, ",", ".") & "','" & Replace(ALIQUTOA_COFINS_N, ",", ".") & "'" & _
               ",'" & Replace(COFINSVCOFINS_N, ",", ".") & "','" & Replace(PisqBCProd_A, ",", ".") & "'" & _
               ",'" & Replace(PisvAliqProd_A, ",", ".") & "','" & Replace(COFINSqBCProd_A, ",", ".") & "','" & Replace(COFINSvAliqProd_A, ",", ".") & "'"

      CONECTA_GLOBAL.Execute "EXEC spMFi010Global " & SQL

      TabItem.MoveNext
   Wend
   If TabItem.State = 1 Then _
      TabItem.Close
   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   PEDIDOitem_INTEGRA_MFI010 = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "PEDIDOitem_INTEGRA_MFI010"
End Function
  
Sub FINANCEIRO_INTEGRA(E1_PREFIXO_A As String, Numr_Nota_A As String)
'On Error GoTo ERRO_TRATA

   Dim Numr_Nota_N As Long

   Numr_Nota_N = 0 & Trim(Numr_Nota_A)

   If Numr_Nota_N <= 0 Then _
      Exit Sub

   If CONECTA_GLOBAL.State <> 1 Then _
      ABRE_BANCO_GLOBAL

   Dim TabPedidoIntegra As New ADODB.Recordset
   Dim TabCabecaIntegra As New ADODB.Recordset
   Dim TabFinac         As New ADODB.Recordset
   Dim strSQL           As String
   Dim PARCELA_N        As Long
   Dim E1_NOMCLI        As String
   Dim E1_EMISSAO       As String
   Dim E1_VENCTO        As String
   Dim E1_CLIENTE       As String
   Dim E1_CARTAO        As String
   Dim E1_ADM           As String
   Dim E1_CARTAUT       As String
   Dim E1_TIPO          As String
   Dim E1_FILIAL_A        As String
   Dim E1_NUM           As String
   Dim E1_NUMNOTA_A       As String
   Dim E1_PARCELA       As String
   Dim E1_LOJA_A          As String
   Dim E1_VENCREA       As String
   Dim E1_VALOR         As String
   Dim E1_BAIXA         As String
   Dim E1_DATABOR       As String
   Dim E1_LA            As String
   Dim E1_MOVIMEN       As String
   Dim E1_SITUACA       As String
   Dim E1_SALDO         As String
   Dim E1_DESCONT       As String
   Dim E1_VALLIQ        As String
   Dim E1_VENCORI       As String
   Dim E1_VLCRUZ        As String
   Dim E1_SERIE         As String
   Dim E1_STATUS        As String
   Dim E1_FLUXO         As String
   Dim E1_DTACRED       As String
   Dim E1_NUMCRD        As String
   Dim E1_FLAG          As String
   Dim E1_OCORREN       As String
   Dim E1_MULTNAT       As String
   Dim E1_PROJPMS       As String
   Dim E1_DESDOBR       As String
   Dim E1_MODSPB        As String

   E1_FILIAL_A = "0" & ESTABELECIMENTO_ID_N
   'E1_LOJA_A = "0" & EMPRESA_ID_N
   E1_LOJA_A = "01"

   E1_MULTNAT = "N"
   E1_PROJPMS = "N"
   E1_DESDOBR = "N"
   E1_MODSPB = "1"

   If Trim(E1_PREFIXO_A) = "" Then
      E1_PREFIXO_A = "NFE"
      Else
         If Trim(E1_PREFIXO_A) = "55" Then _
            E1_PREFIXO_A = "NFE"
         If Trim(E1_PREFIXO_A) = "65" Then _
            E1_PREFIXO_A = "NFC"
   End If

   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close

   strSQL = "select NF.NF_ID, NF.PESSOA_ID, NF.TRANSP_ID, NF.NF_TIPO, NF.NUMR_NOTA, "
   strSQL = strSQL & " NF.SERIE_NOTA, NF.MODELO_DOC, NF.DT_EMISSAO, NF.STATUS, NF.ESTABELECIMENTO_ID, "
   strSQL = strSQL & " PESSOA.DESCRICAO , PESSOA.RAZAO, PESSOA.CNPJCPF "
   strSQL = strSQL & " from NF WITH (NOLOCK)"
   strSQL = strSQL & " INNER JOIN PESSOA WITH (NOLOCK)"
   strSQL = strSQL & " ON NF.PESSOA_ID = PESSOA.PESSOA_ID"

   strSQL = strSQL & " where numr_nota = " & Numr_Nota_N
   strSQL = strSQL & " and modelo_doc = '" & E1_PREFIXO_A & "'"
   strSQL = strSQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabPedidoIntegra.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoIntegra.EOF Then
      CNPJ_CPF_A = "" & TabPedidoIntegra.Fields("cnpjcpf").Value
      E1_NOMCLI = "" & Trim(Left(TabPedidoIntegra.Fields("descricao").Value, 60))
      E1_NOMCLI = "" & Replace(E1_NOMCLI, "'", " ")
      E1_NOMCLI = "" & Replace(E1_NOMCLI, ",", ";")
      Numr_Nota_N = TabPedidoIntegra.Fields("numr_nota").Value
      E1_NUMNOTA_A = Numr_Nota_N
      E1_NUM = Numr_Nota_N
      E1_EMISSAO = "" & TabPedidoIntegra.Fields("dt_emissao").Value
      E1_MOVIMEN = "" & TabPedidoIntegra.Fields("dt_emissao").Value
      E1_DTACRED = "" & TabPedidoIntegra.Fields("dt_emissao").Value
      E1_VENCREA = "" & TabPedidoIntegra.Fields("dt_emissao").Value
      E1_CLIENTE = "" & TRAZ_ID_TABELA_GLOBAL("SA1010", "A1_COD", "A1_CGC", TabPedidoIntegra.Fields("CNPJCPF").Value)
      E1_SITUACA = "0"
      E1_SERIE = "1"
      E1_STATUS = "A"
      E1_FLUXO = "S"
      E1_FLAG = ""
      E1_OCORREN = "01"

'set criar tabela para tratar taxa entrega
      If TabFinac.State = 1 Then _
         TabFinac.Close

      SQL = "select formapagto_id from FORMAPAGTO WITH (NOLOCK)"
      SQL = SQL & " where descricao = 'TAXA ENTREGA'"
      TabFinac.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabFinac.EOF Then _
         FORMAPAGTO_ID_N = 0 & TabFinac.Fields(0).Value


      SQL = "delete SE1010 "
      SQL = SQL & " where E1_NUMNOTA = '" & Trim(E1_NUMNOTA_A) & "'"
      SQL = SQL & " and E1_FILIAL = '" & Trim(E1_FILIAL_A) & "'"
      SQL = SQL & " and E1_LOJA = '" & Trim(E1_LOJA_A) & "'"
      SQL = SQL & " and E1_PREFIXO = '" & Trim(E1_PREFIXO_A) & "'"
      CONECTA_GLOBAL.Execute SQL


      If TabFinac.State = 1 Then _
         TabFinac.Close

      SQL = "select ITEMLANCAMENTO.SEQ , ITEMLANCAMENTO.FORMAPAGTO_ID, ITEMLANCAMENTO.Valor_Item,dt_baixa, "
      SQL = SQL & " ITEMLANCAMENTO.VALOR_DESCONTO,ITEMLANCAMENTO.NUMR_DP , ITEMLANCAMENTO.DT_VENCIMENTO"
      SQL = SQL & " FROM LANCAMENTO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK)"
      SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"

      ''SQL = SQL & " where LANCAMENTO.NUMR_DOC = " & Numr_Nota_N
      SQL = SQL & " where LANCAMENTO.NUMR_DOC = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
      'SQL = SQL & " and FORMAPAGTO_ID <> " & FORMAPAGTO_ID_N

'Debug.Print SQL

      TabFinac.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabFinac.EOF
         PARCELA_N = 0 & TabFinac.Fields("seq").Value
         E1_PARCELA = PARCELA_N
         E1_VENCTO = "" & TabFinac.Fields("dt_vencimento").Value
         E1_VENCORI = "" & TabFinac.Fields("dt_vencimento").Value
         E1_BAIXA = "" & TabFinac.Fields("dt_baixa").Value
         E1_DATABOR = "NULL"
         VALOR_DESCONTO_N = 0 & TabFinac.Fields("valor_desconto").Value
         E1_DESCONT = VALOR_DESCONTO_N
         VALOR_ITEM_N = 0 & (TabFinac.Fields("valor_item").Value - VALOR_DESCONTO_N)
         E1_VALLIQ = "0"
         E1_SALDO = 0   'VALOR_ITEM_N  'yuri pediu pra gravar o troco ou zero
         E1_VALOR = VALOR_ITEM_N
         E1_VLCRUZ = VALOR_ITEM_N
         E1_CARTAO = ""
         E1_ADM = ""
         E1_CARTAUT = ""
         E1_LA = "N"
         E1_NUMCRD = ""

'Campo E1_TIPO=Forma de Pagamento gravar os seguintes numeros :
'01=Dinheiro
'02=Cheque
'03=Cartão de Crédito
'04=Cartão de Débito
'05=Crédito Loja
'10=Vale Alimentação
'11=Vale Refeição
'12=Vale Presente
'13=Vale Combustível
'14=Duplicata Mercantil
'90= Sem pagamento
'99=Outros
'onde os numeros informados 03 e 04 e obrigatorios preencher os campos criados no item 03.
'ver relação pois essas formas dependem de cada empresa

         E1_TIPO = "01"

         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("dinheiro") Then _
            E1_TIPO = "01"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("a vista") Then _
            E1_TIPO = "01"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("á vista") Then _
            E1_TIPO = "01"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("cheque") Then _
            E1_TIPO = "02"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartão de Crédito") Then _
            E1_TIPO = "03"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartão de Credito") Then _
            E1_TIPO = "03"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartao de Crédito") Then _
            E1_TIPO = "03"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartao de Credito") Then _
            E1_TIPO = "03"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartao de Debito") Then _
            E1_TIPO = "04"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartão de Debito") Then _
            E1_TIPO = "04"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartao de Débito") Then _
            E1_TIPO = "04"
         If Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value))) = UCase("Cartão de Débito") Then _
            E1_TIPO = "04"

         If TabFinac.Fields("FORMAPAGTO_ID").Value = 1 Then _
            E1_TIPO = "01"

'MsgBox Trim(UCase(TRAZ_DESCRICAO_FORMAPAGTO(TabFinac.Fields("FORMAPAGTO_ID").Value)))

         If TabConsulta.State = 1 Then _
            TabConsulta.Close
         SQL = "select CARTAOPEDIDO_ID,PEDIDO_ID,BANDEIRA_ID,CNPJ_CARTAO,NUMR_AUTORIZACAO "
         SQL = SQL & " from CARTAOPEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabConsulta.EOF Then
            If Not IsNull(TabConsulta.Fields("CNPJ_CARTAO").Value) Then _
               E1_CARTAO = "" & Trim(TabConsulta.Fields("CNPJ_CARTAO").Value)
            If Not IsNull(TabConsulta.Fields("BANDEIRA_ID").Value) Then _
               E1_ADM = "" & Trim(TabConsulta.Fields("BANDEIRA_ID").Value)
            If Not IsNull(TabConsulta.Fields("NUMR_AUTORIZACAO").Value) Then _
               E1_CARTAUT = "" & Trim(TabConsulta.Fields("NUMR_AUTORIZACAO").Value)
            If Trim(E1_CARTAUT) = "" Then _
               E1_CARTAUT = "000000"
         End If

         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close
         SQL = "select e1_num,R_E_C_N_O_ from SE1010 WITH (NOLOCK)"

         SQL = SQL & " where E1_NUMNOTA = " & Numr_Nota_N
         SQL = SQL & " and e1_parcela = " & PARCELA_N
         SQL = SQL & " and e1_prefixo = '" & E1_PREFIXO_A & "'"

         SQL = SQL & " and E1_FILIAL = '" & E1_FILIAL_A & "'"
         SQL = SQL & " and E1_LOJA = '" & E1_LOJA_A & "'"

         TabCabecaIntegra.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If TabCabecaIntegra.EOF Then
            Acao_N = 1
            Else: Acao_N = 2
         End If
         If TabCabecaIntegra.State = 1 Then _
            TabCabecaIntegra.Close

         SQL = "spSE1010Global " & Acao_N & ",'" & _
                                   E1_PREFIXO_A & "','" & E1_NUM & "','" & _
                                   E1_PARCELA & "','" & E1_TIPO & "','" & E1_CLIENTE & "','" & _
                                   E1_LOJA_A & "','" & E1_NOMCLI & "','" & E1_EMISSAO & "','" & _
                                   E1_VENCTO & "','" & E1_VENCREA & "','" & tpMOEDA(E1_VALOR) & "','" & _
                                   E1_BAIXA & "','" & Null & "','" & E1_LA & "','" & _
                                   E1_MOVIMEN & "','" & E1_SITUACA & "','" & tpMOEDA(E1_SALDO) & "','" & _
                                   tpMOEDA(E1_DESCONT) & "','" & tpMOEDA(E1_VALLIQ) & "','" & E1_VENCORI & "','" & _
                                   tpMOEDA(E1_VLCRUZ) & "','" & E1_NUMNOTA_A & "','" & E1_SERIE & "','" & _
                                   E1_STATUS & "','" & E1_FLUXO & "','" & E1_CARTAO & "','" & _
                                   E1_DTACRED & "','" & E1_NUMCRD & "','" & _
                                   E1_FLAG & "','" & E1_CARTAUT & "','" & E1_FILIAL_A & "','" & _
                                   E1_OCORREN & "','" & E1_MULTNAT & "','" & E1_PROJPMS & "','" & E1_DESDOBR & "','" & E1_MODSPB & "'"

         CONECTA_GLOBAL.Execute "EXEC " & SQL

         TabFinac.MoveNext
         strSQL = ""
      Wend
      If TabFinac.State = 1 Then _
         TabFinac.Close
   End If
   If TabCabecaIntegra.State = 1 Then _
      TabCabecaIntegra.Close
   If TabPedidoIntegra.State = 1 Then _
      TabPedidoIntegra.Close
   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "FINANCEIRO_INTEGRA"
End Sub

Public Sub REIMPRESSAO_NFEC(Variavel_Nome_NFe_XML As String, MES_A As String, ANO_A As String)
'On Error GoTo ERRO_TRATA

   'Shell "UniDANFE.exe a=c:\FalcaoNfe\xml\enviado\200903\31090309252646000130550010000070860000008450-nfe.xml C = ”RETRATO”"

   If Trim(Variavel_Nome_NFe_XML) <> "" Then
      If CONECTA_GLOBAL.State <> 1 Then _
         ABRE_BANCO_GLOBAL

      Dim TabCaminho As New ADODB.Recordset

      If TabCaminho.State = 1 Then _
         TabCaminho.Close

      SQL = "select envionfe from empres"
      TabCaminho.Open SQL, CONECTA_GLOBAL, , , adCmdText
      If Not TabCaminho.EOF Then
         If Not IsNull(TabCaminho.Fields(0).Value) Then

            'ler CAMINHO
            Dim CAMINHO_A As String
            CAMINHO_A = Trim(TabCaminho.Fields(0).Value) & "\Enviado\Autorizados\" & Trim(ANO_A) & "\" & Trim(MES_A) & "\"

            Variavel_Nome_NFe_XML = Trim(Variavel_Nome_NFe_XML)
         End If
      End If

      If TabCaminho.State = 1 Then _
         TabCaminho.Close

      Shell "UniDANFE.exe a=" & Variavel_Nome_NFe_XML & " C = ”RETRATO”"
   End If
   'Shell ("C:\unimake\uninfe\UniDANFE.exe a=" & Variavel_Nome_NFe_XML & " au=" & variavel_com_nome_auxiliar & " c=Paisagem")

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, Me.Name, "REIMPRESSAO_NFEC"
End Sub

Sub PIS_MEGASIM_GLOBAL()
         'Select Case CST_PIS_A
         '   Case "001"    '21 PIS-Operação Tributável com Alíquota Básica  001
         '      PISCST_N = 21
         '   Case "002"     '22 PIS-Operação Tributável com Alíquota por Unidade de Medida de Produto   002
         '      PISCST_N = 22
         '   Case "003"     '23 PIS-Operação Tributável com Alíquota por Unidade de Medida de Produto   003
         '      PISCST_N = 23
         '   Case "004"     '24 PIS-Operação Tributável Monofásica – Revenda a Alíquota Zero   004
         '      PISCST_N = 24
         '   Case "005"     '25 PIS-Operação Tributável por Substituição Tributária   005
         '      PISCST_N = 25
         '   Case "006"     '26 PIS-Operação Tributável a Alíquota Zero   006
         '      PISCST_N = 26
         '   Case "007"     '27 PIS-Operação Isenta da Contribuição 007
         '      PISCST_N = 27
         '   Case "008"     '28 PIS-Operação sem Incidência da Contribuição  008
         '      PISCST_N = 28
         '   Case "009"     '29 PIS-Operação com Suspensão da Contribuição   009
         '      PISCST_N = 29
         '   Case "049"     '30 PIS-Outras Operações de Saída 049
         '      PISCST_N = 30
         '   Case "050"     '31 PIS-Operação com Direito a Crédito – Vinculada Exclusivamente a Receita Tributada no Mercado Interno  050
         '      PISCST_N = 31
         '   Case "098"     '52 PIS-Outras Operações de Entrada  098
         '      PISCST_N = 52
         '   Case "099"     '53 PIS-Outras Operações 099
         '      PISCST_N = 53
         'End Select
End Sub

Sub COFINS_MEGASIM_GLOBAL()
         'Select Case CST_COFINS_A
         '   Case "001"     '54 COFINS-Operação Tributável com Alíquota Básica 001
         '      COFINSCST_N = 54
         '   Case "002"     '55 COFINS-Operação Tributável com Alíquota Diferenciada 002
         '      COFINSCST_N = 55
         '   Case "003"     '56 COFINS-Operação Tributável com Alíquota por Unidade de Medida de Produto 003
         '      COFINSCST_N = 56
         '   Case "004"     '57 COFINS-Operação Tributável Monofásica - Revenda a Alíquota Zero 004
         '      COFINSCST_N = 57
         '   Case "005"     '58 COFINS-Operação Tributável por Substituição Tributária 005
         '      COFINSCST_N = 58
         '   Case "006"     '59 COFINS-Operação Tributável a Alíquota zero 006
         '      COFINSCST_N = 59
         '   Case "007"     '60 COFINS-Operação Isenta da Contribuição 007
         '      COFINSCST_N = 60
         '   Case "008"     '61 COFINS-Operação sem Incidência da Contribuição 008
         '      COFINSCST_N = 61
         '   Case "009"     '62 COFINS-Operação com Suspensão da Contribuição 009
         '      COFINSCST_N = 52
         '   Case "049"     '63 COFINS-Outras Operações de Saída 049
         '      COFINSCST_N = 63
         '   Case "098"     '85 COFINS-Outras Operações de Entrada 098
         '      COFINSCST_N = 85
         '   Case "099"     '86 COFINS-Outras Operações 099
         '      COFINSCST_N = 86
         'End Select
End Sub
