Attribute VB_Name = "mdlOS"
Public Sub ABRE_BANCO_SQLSERVER(Base_Dados As String)
On Error GoTo ERRO_TRATA

   Dim ConnectString As String

   If CONECTA_RETAGUARDA.State <> 1 Then
      SENHA_ADM_SQLSERVER = "ejsnenas"
      USUARIO_ADM_SQLSERVER = "sa"

      ConnectString = "uid=" & USUARIO_ADM_SQLSERVER & _
                      ";pwd=" & SENHA_ADM_SQLSERVER & _
                      ";Provider=SQLOLEDB.1;Server=" & SERVIDOR_MEGASIM & _
                      ";database=" & Base_Dados & _
                      ";dsn='" & SERVIDOR_MEGASIM & _
                      "';connection=adConnectAsync"

      With CONECTA_RETAGUARDA
         .ConnectionString = ConnectString
         .ConnectionTimeout = 10
         .Open
      End With

      Crystaldsn = SERVIDOR_MEGASIM
      Crystaldsq = Base_Dados
      Crystaluid = USUARIO_ADM_SQLSERVER
      Crystalpwd = SENHA_ADM_SQLSERVER

      AUTENTICA_GRID = "Provider=SQLOLEDB.1;Password='" & SENHA_ADM_SQLSERVER & "'"
      AUTENTICA_GRID = AUTENTICA_GRID & ";Persist Security Info=True;User ID='" & USUARIO_ADM_SQLSERVER & "'"
      AUTENTICA_GRID = AUTENTICA_GRID & ";Initial Catalog=" & Base_Dados
      AUTENTICA_GRID = AUTENTICA_GRID & ";Data Source=" & SERVIDOR_MEGASIM
   End If

   strFormatacao2Digitos = "##,##0.00"
   strFormatacao3Digitos = "##,##0.000"
   strFormatacao4Digitos = "##,##0.0000"
   strFormatacao5Digitos = "##,##0.00000"
   strFormatacao6Digitos = "##,##0.000000"
   strFormatacao7Digitos = "##,##0.0000000"
   strFormatacao8Digitos = "##,##0.00000000"

Exit Sub
ERRO_TRATA:
   MsgBox "Erro ao abrir banco de dados : " & Err.Description
   End
End Sub

Public Sub ABRE_BANCO_GLOBAL()
'On Error GoTo ERRO_TRATA

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   If Not FSO.FileExists(App.Path & "\GLOBAL.ini") Then
      MsgBox "Arquivo de inicializa��o do sistema n�o encontrado, entre em contato com suporte."
      End
   End If

   Dim Usuario_B        As String
   Dim Senha_B          As String
   Dim Nome_Banco       As String

   f = FreeFile

   Open App.Path & "\GLOBAL.INI" For Input As f

   Line Input #f, sLine
   Servidor_Global = sLine

   Line Input #f, sLine
   Nome_Banco = sLine

   Close #f

   Dim ConnectString As String

   Senha_B = "ejsnenas"
   Usuario_B = "sa"

   ConnectString = "uid=" & Usuario_B & _
                   ";pwd=" & Senha_B & _
                   ";Provider=SQLOLEDB.1;Server=" & Servidor_Global & _
                   ";database=" & Nome_Banco & _
                   ";dsn='" & Servidor_Global & _
                   "';connection=adConnectAsync"

   With CONECTA_GLOBAL
      .ConnectionString = ConnectString
      .ConnectionTimeout = 10
      .Open
   End With

   'Crystaldsn = SERVIDOR_SHFSYS
   'Crystaldsq = nome_banco
   'Crystaluid = usuario_b
   'Crystalpwd = senha_b

   'Autentica_GRID = "Provider=SQLOLEDB.1;Password='" & senha_b & "'"
   'Autentica_GRID = Autentica_GRID & ";Persist Security Info=True;User ID='" & usuario_b & "'"
   'Autentica_GRID = Autentica_GRID & ";Initial Catalog=" & nome_banco
   'Autentica_GRID = Autentica_GRID & ";Data Source=" & SERVIDOR_SHFSYS

Exit Sub
ERRO_TRATA:
   MsgBox "Erro ao abrir banco GLOBAL"
End Sub
Public Sub ABRE_BANCO_AUXILIAR(Base_Dados As String, Servidor_Dados As String)
'On Error GoTo ERRO_TRATA

   Dim ConnectString As String

   If CONECTA_AUXILIAR.State <> 1 Then
      SENHA_ADM_SQLSERVER = "ejsnenas"
      USUARIO_ADM_SQLSERVER = "sa"

      ConnectString = "uid=" & USUARIO_ADM_SQLSERVER & _
                      ";pwd=" & SENHA_ADM_SQLSERVER & _
                      ";Provider=SQLOLEDB.1;Server=" & Servidor_Dados & _
                      ";database=" & Base_Dados & _
                      ";dsn='" & Servidor_Dados & _
                      "';connection=adConnectAsync"

      With CONECTA_AUXILIAR
         .ConnectionString = ConnectString
         .ConnectionTimeout = 10
         .Open
      End With
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "ABRE_BANCO_AUXILIAR"
End Sub

Public Sub INICIALIZA_SISTEMA()
'On Error GoTo ERRO_TRATA

   INDR_LOGON = False

   If Not FSO.FileExists(App.Path & "\MEGASIM.ini") Then
   'If Not FSO.FileExists("c:\megasim\MEGASIM.ini") Then
      MsgBox "Arquivo de inicializa��o do sistema n�o encontrado, entre em contato com suporte."
      End
   End If

   f = FreeFile

   Open App.Path & "\MEGASIM.ini" For Input As f
   'Open "c:\megasim\MEGASIM.ini" For Input As f

   Line Input #f, sLine
   SERVIDOR_MEGASIM = sLine

   Line Input #f, sLine
   NOME_BANCO_DADOS = sLine

   Line Input #f, sLine
   PATH_REL = sLine

   Line Input #f, sLine
   PATH_TXT = sLine

   Line Input #f, sLine
   EMPRESA_ID_N = sLine

   Line Input #f, sLine
   INDR_LOGON = sLine

   Line Input #f, sLine
   INDR_CAIXA = sLine

   Line Input #f, sLine
   NOME_ESTABELEC = sLine

   Line Input #f, sLine
   INDR_TREINAMENTO = sLine

   Line Input #f, sLine
   NUMERO_CAIXA_CPU = sLine

   Line Input #f, sLine
   ESTABELECIMENTO_ID_N = sLine

   Line Input #f, sLine
   INDR_REMOTO = sLine

   Line Input #f, sLine
   USA_TEF = sLine

'   Line Input #f, sLine
   INDR_OS_VEICULO = False

   Close #f

   If Trim(Command) <> "" Then _
      NOME_BANCO_DADOS = Trim(Command)
   '   Else
   'If Trim(NOME_BANCO_DADOS) = "" Then _
      NOME_BANCO_DADOS = Trim(InputBox("Informe BANCO DE DADOS PARA CONEX�O", "SHF INFORM�TICA", NOME_BANCO_DADOS))

If INDR_PANIFICADORA = True Then
   If INDR_CAIXA = True Then
      If (App.PrevInstance) Then
          Dim nome_tela As String
          nome_tela = App.Title
          App.Title = "Sistama j� est� aberto, verifique !!!"
          'AppActivate  nome_tela
          SendKeys "%R", True
          MsgBox "Sistama j� est� aberto, verifique !!!"
          End
          Exit Sub
      End If
   End If
End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "INICIALIZA_SISTEMA"
End Sub

Public Sub TRATA_ERROS(Desc_Erro As Variant, Formulario As String, Objeto As String)
   SQL3 = "Porfavor, descreva detalhes do(s) erro(s) : " & Err.Number & " - " & Desc_Erro
   Select Case Err.Number
      'Erros gen�ricos.
      Case 3005  'Database Name ' isn't a valid database name.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3006  'Database 'name' is exclusively locked.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3008  'Table 'name' is exclusively locked.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3009  'Couldn 't lock table 'name'; currently in use.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3010  'Table 'name' already exists.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3015  'Index name' isn't an index in this table.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3019  'Operation invalid without a current index.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3020  'Update or CancelUpdate without AddNew or Edit.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3021  'No current record.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3022  'Duplicate value in index, primary key, or relationship. Changes were unsuccessful.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3023  'AddNew or Edit already used.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3034  'Commit or Rollback without BeginTrans.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3036  'Database has reached maximum size.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3037  'can 't open any more tables or queries.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3040  'Disk I/O error during read.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3044  'Path' isn't a valid path.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3046  'Couldn 't save; currently locked by another user.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub

      'Erros relacionados a bloqueio de registros
      Case 3027  'can 't update.  Database or object is read-only.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3158  'Couldn 't save record; currently locked by another user.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3167  'Record is deleted.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3186  'Couldn 't save; currently locked by user 'name' on machine 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3187  'Couldn 't read; currently locked by user 'name' on machine 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3188  'Couldn 't update; currently locked by another session on this machine.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3189  'Table 'name' is exclusively locked by user 'name' on machine 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3197  'Data has changed; operation stopped.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3260  'Couldn 't update; currently locked by user 'name' on machine 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3261  'Table 'name' is exclusively locked by user 'name' on machine 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3356  'The database is opened by user 'name' on machine 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub

      'Erros relacionados a Permiss�es
      Case 3107  'Record(s) can't be added; no Insert Data permission on 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3108  'Record(s) can't be edited; no Update Data permission on 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3109  'Record(s) can't be deleted; no Delete Data permission on 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3110  'Couldn 't read definitions; no Read Definitions permission for table or query 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3111  'Couldn 't create; no Create permission for table or query 'name'.
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3112  'Record(s) can't be read; no Read Data permission on 'name'.   End Select
         CRITERIO = InputBox(SQL3, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
   End Select

   CRITERIO = InputBox(SQL3 & " / " & Objeto, "Formul�rio: " & Formulario & "; Rotina:" & Objeto)
   
   GRAVA_ERRO Formulario, Objeto
   'Exit Sub

'Resume Next Retorna a execu��o na linha que vem logo ap�s � linha que gerou o erro
'Resume Executa mais uma vez a linha que gerou o erro.
'Resume Label Retorna a execu��o da linha que vem ap�s a etiqueta citada.
'Resume Number Retorna a execu��o na linha com o n�mero indicado.
'Exit Sub Sai da sub rotina atual.
'Exit Function Sai da fun��o atual.
'Exit Property Sai da propriedade atual.
'On Error Redefine a l�gica de tratamento de erros.
'Err.Clear Elimina o erro sem afetar a execu��o do programa.
'End Encerra a execu��o do programa.
'Number Fornece Numero do erro gerado
'Description Fornece a descri��o do erro.
'Source Identifica o nome do objeto que gerou o erro
'Raise Gera um erro de execu��o, usado para testar condi��es de erro.
End Sub

Public Sub GRAVA_ERRO(Formulario As String, Objeto As String)
   If CONECTA_RETAGUARDA.State = 1 Then
      SqL2 = Replace(Err.Description, ",", " ")
      SqL2 = Replace(SqL2, "'", " ")

      CRITERIO = Replace(CRITERIO, ",", " ")
      CRITERIO = Replace(CRITERIO, "'", " ")

      SQL = "insert into ERRO values("
      SQL = SQL & MAX_ID("erro" & "_id", "ERRO", "", "", "", "")
      SQL = SQL & "," & Err.Number
      SQL = SQL & ",'" & Replace(SqL2, "',", " ") & "'"
      SQL = SQL & ",'" & CRITERIO & "'"
      SQL = SQL & ",'" & "FORMUL�RIO: " & Formulario & "   -   OBJETO: " & Objeto & "'"
      SQL = SQL & ",'" & DMA(Date) & " " & Time & "'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      CRITERIO = ""
      SqL2 = ""
      Err.Clear
   End If
End Sub

Public Sub PERGUNTA(Msg As String, Style As Variant, Title As String, Help As String, Ctxt As Long)
'On Error GoTo ERRO_TRATA

   Beep

   RESPOSTA = MsgBox(Msg, Style, Title, Help, Ctxt)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "PERGUNTA"
End Sub

Public Function CALCULACNPJ(Numero As String) As String 'Fun��es para Validar CPF e CNPJ Validar CNPJ
'On Error GoTo ERRO_TRATA

   Dim i As Integer
   Dim prod As Integer
   Dim mult As Integer
   Dim Digito As Integer

   If Not IsNumeric(Numero) Then
      CALCULACNPJ = ""
      Exit Function
   End If

   mult = 2
   For i = Len(Numero) To 1 Step -1
     prod = prod + Val(Mid(Numero, i, 1)) * mult
     mult = IIf(mult = 9, 2, mult + 1)
   Next

   Digito = 11 - Int(prod Mod 11)
   Digito = IIf(Digito = 10 Or Digito = 11, 0, Digito)

   CALCULACNPJ = Trim(Str(Digito))

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "CALCULACNPJ"
End Function

Public Function VALIDACNPJ(CNPJ As String) As Boolean
'On Error GoTo ERRO_TRATA

   If CALCULACNPJ(Left(CNPJ, 12)) <> Mid(CNPJ, 13, 1) Then
      VALIDACNPJ = False
      Exit Function
   End If
   
   If CALCULACNPJ(Left(CNPJ, 13)) <> Mid(CNPJ, 14, 1) Then
      VALIDACNPJ = False
      Exit Function
   End If
   VALIDACNPJ = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "VALIDACNPJ"
End Function

Public Function CHECA_CNPJCPF(CNPJ_CPF_A As String) As Boolean
   CHECA_CNPJCPF = False

   CNPJ_CPF_A = Trim(CNPJ_CPF_A)

   If CInt(Len(CNPJ_CPF_A)) = 11 Then
      CHECA_CNPJCPF = ValidaCPF(CNPJ_CPF_A)
      Else
         If CInt(Len(CNPJ_CPF_A)) = 14 Then
            CHECA_CNPJCPF = VALIDACNPJ(CNPJ_CPF_A)
         End If
   End If
End Function

'2- Validar CPF
Function CALCULACPF(CPF As String) As Boolean
'On Error GoTo ERRO_TRATA

   'Esta rotina foi adaptada da revista F�rum Access
   On Error GoTo Err_CPF

   Dim i As Integer        'utilizada nos FOR... NEXT
   Dim strCampo As String  'armazena do CPF que ser� utilizada para o c�lculo
   Dim strCaracter As String   'armazena os d�gitos do CPF da direita para a esquerda
   Dim intNumero As Integer    'armazena o digito separado para c�lculo (uma a um)
   Dim intMais As Integer  'armazena o digito espec�fico multiplicado pela sua base
   Dim lngSoma As Long     'armazena a soma dos d�gitos multiplicados pela sua base(intmais)
   Dim dblDivisao As Double    'armazena a divis�o dos d�gitos * base por 11
   Dim lngInteiro As Long  'armazena inteiro da divis�o
   Dim intResto As Integer     'armazena o resto
   Dim intDig1 As Integer  'armazena o 1� digito verificador
   Dim intDig2 As Integer  'armazena o 2� digito verificador
   Dim strConf As String   'armazena o digito verificador

   lngSoma = 0
   intNumero = 0
   intMais = 0
   strCampo = Left(CPF, 9)

   'Inicia c�lculos do 1� d�gito
   For i = 2 To 10
      strCaracter = Right(strCampo, i - 1)
      intNumero = Left(strCaracter, 1)
      intMais = intNumero * i
      lngSoma = lngSoma + intMais
   Next i
   dblDivisao = lngSoma / 11

   lngInteiro = Int(dblDivisao) * 11
   intResto = lngSoma - lngInteiro
   If intResto = 0 Or intResto = 1 Then
      intDig1 = 0
      Else
         intDig1 = 11 - intResto
   End If

   strCampo = strCampo & intDig1 'concatena o CPF com o primeiro digito verificador
   lngSoma = 0
   intNumero = 0
   intMais = 0
   'Inicia c�lculos do 2� d�gito
   For i = 2 To 11
     strCaracter = Right(strCampo, i - 1)
     intNumero = Left(strCaracter, 1)
     intMais = intNumero * i
     lngSoma = lngSoma + intMais
   Next i
   dblDivisao = lngSoma / 11
   lngInteiro = Int(dblDivisao) * 11
   intResto = lngSoma - lngInteiro
   If intResto = 0 Or intResto = 1 Then
      intDig2 = 0
      Else
         intDig2 = 11 - intResto
   End If
   strConf = intDig1 & intDig2
   'Caso o CPF esteja errado dispara a mensagem
   If strConf <> Right(CPF, 2) Then
      CALCULACPF = False
      Else
         CALCULACPF = True
   End If
   Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "CALCULACPF"
End Function

'====================================================
Public Sub INSERIR_BRANCO(TAMANHO_N As Long, STRING_A As String)
'On Error GoTo ERRO_TRATA

   STRING_A = Trim(STRING_A)
   While Len(STRING_A) < TAMANHO_N
      STRING_A = STRING_A & " "
   Wend
   CRITERIO = ""
   CRITERIO = STRING_A

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "INSERIR_BRANCO"
End Sub
'================================
Public Sub ParseToArray(sLine As String, a() As String)
'On Error GoTo ERRO_TRATA

   Dim p As Long, LastPos As Long, i As Long

   p = InStr(sLine, ";")

   Do While p
      a(i) = Mid$(sLine, LastPos + 1, p - LastPos - 1)
      LastPos = p
      i = i + 1
      p = InStr(LastPos + 1, sLine, ";", vbBinaryCompare)
   Loop
   a(i) = Mid$(sLine, LastPos + 1)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "ParseToArray"
End Sub

Public Sub ParseToArrayVIRGULA(sLine As String, a() As String)
'On Error GoTo ERRO_TRATA

   Dim p As Long, LastPos As Long, i As Long

   sLine = Replace(sLine, "null", "")

   p = InStr(sLine, ",")
   i = 1

   Do While p
      'On Error GoTo SAI_LOOP
      On Error Resume Next
      a(i) = Mid$(sLine, LastPos + 1, p - LastPos - 1)
      LastPos = p
      i = i + 1
      p = InStr(LastPos + 1, sLine, ",", vbBinaryCompare)
   Loop
   a(i) = Mid$(sLine, LastPos + 1)

SAI_LOOP:
   Err.Clear

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "ParseToArrayVIRGULA"
End Sub

Public Function Exporta(db As ADODB.Recordset, sSQL As String, sDestino As String) As Boolean
'On Error GoTo ERRO_TRATA

   Dim Registro      As New ADODB.Recordset
   Dim nI            As Long
   Dim nJ            As Long
   Dim nArquivo      As Integer
   Dim sTemp         As String

   Registro.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   'Abre o arquivo de saida
   nArquivo = FreeFile

   Open sDestino For Output As #nArquivo

   'Exporta os nomes dos campos
   For nI = 0 To Registro.Fields.Count - 1
      sTemp = (Registro.Fields(nI).Name)
      Write #nArquivo, sTemp;
   Next
   Write #nArquivo,

   If Registro.RecordCount > 0 Then
      Registro.MoveLast
      Registro.MoveFirst

      For nI = 1 To Registro.RecordCount
         For nJ = 0 To Registro.Fields.Count - 1
            If Not IsNull(Registro.Fields(nJ)) Then
               sTemp = Replace(Registro.Fields(nJ), ",", ".")
               sTemp = Replace(sTemp, """", "")
               Else: sTemp = "null"
            End If
            Write #nArquivo, Trim(sTemp);
         Next
         Write #nArquivo,
         Registro.MoveNext
      Next
   End If

   Close #nArquivo
   Exporta = True
   
   Exit Function

ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "Exporta"
   Exporta = False
End Function

Public Sub LoadEXE(Dir As String)
'On Error GoTo ERRO_TRATA

   Dim X As Integer
   Dim nofreeze As Integer
   
   X = Shell(Dir, 1)
   nofreeze = DoEvents()
   Exit Sub

   'mensagem de erro personalizada
   If Err.Number = 6 Then Exit Sub
   MsgBox "O aplicativo n�o foi localizado !!! Verifique sua localiza��o ...", vbExclamation
   'Se preferir use a mensagem de erro padrao
   'MsgBox "Error:" & vbCrLf & err.Description & vbCrLf & err.Number, vbExclamation

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "LoadEXE"
'   Resume
End Sub

Public Sub CentralizaJanela(Form As Form)
'On Error GoTo ERRO_TRATA

    Form.Top = (Screen.Height - Form.Height) / 2
    Form.Left = (Screen.Width - Form.Width) / 2
    Form.Top = 585

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "CentralizaJanela"
End Sub

Public Function MAX_ID(Nome_Campo_Max As String, Nome_Tabela, Campo_01 As String, Info_01 As String, Campo_02 As String, Info_02 As String) As Long
'On Error GoTo ERRO_TRATA

   Dim TabID   As New ADODB.Recordset
   Dim strsql  As String
 
   If TabID.State = 1 Then _
      TabID.Close
 
   MAX_ID = 0

   strsql = "select max(" & Nome_Campo_Max & ") from " & Nome_Tabela
   strsql = strsql & " where " & Nome_Campo_Max & " is not null "
   If (Trim(Campo_01) <> "") And (Info_01 <> "") Then _
      strsql = strsql & " and " & Trim(Campo_01) & " = " & Trim(Info_01)
   If (Trim(Campo_02) <> "") And (Info_02 <> "") Then _
      strsql = strsql & " and " & Trim(Campo_02) & " = " & Trim(Info_02)
   TabID.Open strsql, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabID.EOF Then _
      If Not IsNull(TabID.Fields(0).Value) Then _
         MAX_ID = TabID.Fields(0).Value + 1

   If TabID.State = 1 Then _
      TabID.Close

   If MAX_ID <= 0 Then _
      MAX_ID = 1

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "MAX_ID"
End Function
'============== por extenso
Public Function Extenso(nvalor)
'On Error GoTo ERRO_TRATA

    'Valida Argumento
   If IsNull(nvalor) Or nvalor <= 0 Or nvalor > 9999999.99 Then _
      Exit Function

    'Vari�veis
    Dim nContador, nTamanho As Integer
    Dim cValor, cParte, cFinal As String
    ReDim aGrupo(4), aTexto(4) As String
    
    'Matrizes de extensos (Parciais)
    ReDim aUnid(19) As String
    aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
    aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
    aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
    aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
    aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
    aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
    aUnid(19) = "dezenove "
    
    ReDim aDezena(9) As String
    aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
    aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
    aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
    aDezena(9) = "noventa "
    
    ReDim aCentena(9) As String
    aCentena(1) = "cento ": aCentena(2) = "duzentos "
    aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
    aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
    aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
    aCentena(9) = "novecentos "
    
    'Separa valor em grupos
    cValor = Format$(nvalor, "0000000000.00")
    aGrupo(1) = Mid$(cValor, 2, 3)
    aGrupo(2) = Mid$(cValor, 5, 3)
    aGrupo(3) = Mid$(cValor, 8, 3)
    aGrupo(4) = "0" + Mid$(cValor, 12, 2)
    
    'Calcula cada grupo
    For nContador = 1 To 4
      cParte = aGrupo(nContador)
      nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
      If nTamanho = 3 Then
        If Right$(cParte, strFormatacao2Digitos) <> "00" Then
          aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
          nTamanho = 2
        Else
          aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", _
          aCentena(Left(cParte, 1)))
        End If
      End If
      If nTamanho = 2 Then
        If Val(Right(cParte, 2)) < 20 Then
          aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
        Else
          aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
          If Right$(cParte, 1) <> "0" Then
            aTexto(nContador) = aTexto(nContador) + "e "
            nTamanho = 1
          End If
        End If
      End If
      If nTamanho = 1 Then
        aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
      End If
    Next
    
    'Final
    If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
      cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos")
    Else
      cFinal = ""
      cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, _
      "milh�es ", "milh�o "), "")
      If Val(aGrupo(2) + aGrupo(3)) = 0 Then
        cFinal = cFinal + "de "
      Else
        cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "mil ", "")
      End If
      cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "real ", _
      "reais ")
      cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, _
     "centavo", "centavos"), "")
    End If
    Extenso = UCase$(cFinal)

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "Extenso"
End Function

Public Sub INSERIR_0(TAMANHO_N As Long, STRING_A As String)
'On Error GoTo ERRO_TRATA

   STRING_A = Trim(STRING_A)
   While Len(STRING_A) < TAMANHO_N
      STRING_A = "0" & STRING_A
      STRING_A = Trim(STRING_A)
   Wend
   CRITERIO = ""
   CRITERIO = Trim(STRING_A)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "INSERIR_0"
End Sub

Public Sub DestacaTexto(Objeto As TextBox)
'On Error GoTo ERRO_TRATA

    Objeto.SelStart = 0
    Objeto.SelLength = Len(Objeto.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "DestacaTexto"
End Sub

Public Function CHECA_ABERTURA_DIA() As Boolean
   CHECA_ABERTURA_DIA = False

   Dim Ano_A As String
   Dim Mes_A As String
   Dim Dia_A As String

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   SQL = "select caixadia_id,dt_abertura from CAIXADIA "
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and dt_abertura >= '" & DMA(Date) & "'"
   SQL = SQL & " and dt_abertura <= '" & DMA(Date) & "'"
   SQL = SQL & " and tipo = 'B'" 'caixa balc�o
   SQL = SQL & " and ESTABELECIMENTO_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and NUMERO_CAIXA_CPU = " & NUMERO_CAIXA_CPU

   TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCAIXA.EOF Then
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      MsgBox "Caixa do dia n�o foi aberto. Efetuar abertura do caixa."
      'ABRE_CAIXA
      Exit Function
      Else
         If TabEmpresa.State = 1 Then _
            TabEmpresa.Close

         SQL = "SELECT par FROM EMPRESA "
         SQL = SQL & " INNER JOIN ESTABELECIMENTO "
         SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
         SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabEmpresa.EOF Then
            If Not IsNull(TabEmpresa.Fields(0).Value) Then

               If Trim(TabEmpresa.Fields(0).Value) = "" Then
                  MsgBox "Solicitar Contra Senha ao suporte da SHF INFORM�TICA."
                  GRAVA_PRIMEIRA_DATA Date
               End If

               SqL2 = TabEmpresa.Fields(0).Value

               frmCRIPTO.txtCripto.Text = SqL2
               Call frmCRIPTO.DECODIFICA

               SqL2 = frmCRIPTO.txtDeCripto.Text

               SQL = Left(SqL2, 2)
               SQL = SQL & "/" & Mid(SqL2, 7, 2)
               SQL = SQL & "/" & Mid(SqL2, 3, 4)

               DATA_INI = DMA(SQL)

               Else
                  MsgBox "Solicitar Contra Senha ao suporte da SHF INFORM�TICA."
                  GRAVA_PRIMEIRA_DATA Date
            End If
            Else
               MsgBox "Solicitar Contra Senha ao suporte da SHF INFORM�TICA."
               GRAVA_PRIMEIRA_DATA Date
         End If
         If TabEmpresa.State = 1 Then _
            TabEmpresa.Close

         If Date >= DATA_INI Then
            SqL2 = "" & InputBox("Solicitar contra senha para libera��o do sistema.", "Informe Chave para libera��o.")

            frmCRIPTO.txtCripto.Text = SqL2
            Call frmCRIPTO.DECODIFICA

            SQL3 = frmCRIPTO.txtDeCripto.Text
            DATA_INI = DMA(SQL3)

            If Date >= DATA_INI Then
               MsgBox "Chave informada inv�lida, Solicitar Contra Senha ao suporte da SHF INFORM�TICA. " & SQL3
               Else: GRAVA_PRIMEIRA_DATA SQL3
            End If

            End
         End If

'=============

         'DATA_INI = Format(Date, "dd/mm/yyyy")

         'If DATA_INI >= "10/05/2013" Then
         '   MsgBox "Solicitar Contra Senha ao suporte da SHF INFORM�TICA."
         '   End
         'End If

         If Date <> TabCAIXA.Fields("dt_abertura").Value Then
            MsgBox "Data do sistema incorreta, Atualizar sistema."
            End
         End If
   End If

   If TabCAIXA.State = 1 Then _
      TabCAIXA.Close

   CHECA_ABERTURA_DIA = True
End Function

Public Function Aciona_Tecla(key As Integer)
   If (key = 13) Then
      SendKeys "{tab}"
      Else: tecla = key
   End If
End Function

Public Sub MOSTRA_RODAPE(Msg1 As String, Msg2 As String, Msg3 As String, Msg4 As String, Msg5 As String)
   If Trim(Msg1) <> "" Then
      frmINICIO.BARI.Panels.Clear
      frmINICIO.BARI.Panels.Add (1)
      frmINICIO.BARI.Panels(1).Text = Trim(Msg1)
      frmINICIO.BARI.Panels(1).AutoSize = sbrContents
      If Trim(Msg2) <> "" Then
         frmINICIO.BARI.Panels.Add (2)
         frmINICIO.BARI.Panels(2).Text = Trim(Msg2)
         frmINICIO.BARI.Panels(2).AutoSize = sbrContents
         If Trim(Msg3) <> "" Then
            frmINICIO.BARI.Panels.Add (3)
            frmINICIO.BARI.Panels(3).Text = Trim(Msg3)
            frmINICIO.BARI.Panels(3).AutoSize = sbrContents
            If Trim(Msg4) <> "" Then
               frmINICIO.BARI.Panels.Add (4)
               frmINICIO.BARI.Panels(4).Text = Trim(Msg4)
               frmINICIO.BARI.Panels(4).AutoSize = sbrContents
               If Trim(Msg5) <> "" Then
                  frmINICIO.BARI.Panels.Add (5)
                  frmINICIO.BARI.Panels(5).Text = Trim(Msg5)
                  frmINICIO.BARI.Panels(5).AutoSize = sbrContents
               End If
            End If
         End If
      End If
   End If
End Sub
'=============cores
Sub FadeForm(frm As Form, pRed As Integer, pGreen As Integer, pBlue As Integer)
'On Error GoTo ERRO_TRATA

    Dim SaveScale As Integer, SaveStyle As Integer, SaveDraw As Integer
    Dim Y As Long, X As Long, i As Long, j As Long, pixels As Long
    'salvar as configura��es atuais do form
    SaveScale = frm.ScaleMode
    SaveStyle = frm.DrawStyle
    SaveDraw = frm.AutoRedraw
    'pintar a tela
    frm.ScaleMode = 3
    pixels = Screen.Height / Screen.TwipsPerPixelY
    X = pixels / 64 + 0.5

    frm.DrawStyle = 5
    frm.AutoRedraw = True
    For j = 0 To pixels Step X
        Y = 240 - 245 * j / pixels
        'Y = 240 - 245 * j / pixels
        If Y < 0 Then Y = 0
        frm.Line (-2, j - 2)-(Screen.Width + 2, j + X + 3), RGB(-pRed * Y, -pGreen * 111, -pBlue * Y), BF
    Next j
    'restaura configura��es do form
    frm.ScaleMode = SaveScale
    frm.DrawStyle = SaveStyle
    frm.AutoRedraw = SaveDraw

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, frm.Name, "FadeForm"
End Sub

Public Function Bissexto(intAno As Integer) As Boolean
   Bissexto = False

   If intAno Mod 4 = 0 Then
      If intAno Mod 100 = 0 Then
         If intAno Mod 400 = 0 Then _
            Bissexto = True
         Else: Bissexto = True
      End If
   End If

End Function

'========================================
Public Function Libera_Acesso(ROTINA_LIBERA As String) As Boolean
   Libera_Acesso = False

   If CODG_USU_N <> 144 Then
      Dim TabAcesso As New ADODB.Recordset

      If TabAcesso.State = 1 Then _
         TabAcesso.Close

      SQL = "SELECT PERMISSAO.Menuid, PERMISSAO.Usuid, PERMISSAO.Acesso, USUARIO.USUARIO_ID, "
      SQL = SQL & " USUARIO.PESSOA_ID, USUARIO.EMPRESA_ID, USUARIO.NOME, USUARIO.SENHA, "
      SQL = SQL & " USUARIO.CPF, USUARIO.TIPO, USUARIO.NIVEL, "
      SQL = SQL & " USUARIO.Status, USUARIO.Logon, USUARIO.CLASSE"
      SQL = SQL & " FROM PERMISSAO "
      SQL = SQL & " INNER JOIN USUARIO "
      SQL = SQL & " ON PERMISSAO.Usuid = USUARIO.USUARIO_ID"

      SQL = SQL & " where usuario_id = " & CODG_USU_N
      SQL = SQL & " and Menuid = '" & Trim(ROTINA_LIBERA) & "'"

      TabAcesso.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabAcesso.EOF Then
         Libera_Acesso = True
         Else: Libera_Acesso = False
      End If
      If TabAcesso.State = 1 Then _
         TabAcesso.Close
      Else: Libera_Acesso = True
   End If
   Libera_Acesso = True
End Function

Public Function Verifica_dia(dia, Mes)
   Dim diasDoMes As Variant

   dia = Val(dia)

   diasDoMes = Array(31, 28, 30, 30, 31, 30, 31, 30, 30, 31, 30, 31)

   If dia = 31 Then
      Verifica_dia = diasDoMes(Mes - 1)
      Else: Verifica_dia = dia
   End If
   
End Function
Public Sub SP_PROCURA_VENDEDOR(Codg_Eq_n As Long, _
                               VENDEDOR_ID_N As Long, _
                               Nome_Eq_a As String, _
                               NOME_VEND_A As String, _
                               CGC_CPF As String, _
                               CPF As String, _
                               Status As String)
   SQL = "select * from EQUIPE e "
   SQL = SQL & " inner join VENDEDOR v "
   SQL = SQL & " on e.codg_eq = v.codg_eq "
   SQL = SQL & " where e.empresa_id = " & EMPRESA_ID_N
   If Codg_Eq_n > 0 Then _
      SQL = SQL & " and e.codg_eq = " & Codg_Eq_n
   If VENDEDOR_ID_N > 0 Then _
      SQL = SQL & " and v.vendedor_id = " & VENDEDOR_ID_N
   If Nome_Eq_a <> "" Then _
      SQL = SQL & " and e.descricao = '" & Nome_Eq_a & "'"
   If NOME_VEND_A <> "" Then _
      SQL = SQL & " and v.nome_vend = '" & NOME_VEND_A & "'"
   If CGC_CPF <> "" Then _
      SQL = SQL & " and e.cgccpf = '" & CGC_CPF & "'"
   If CPF <> "" Then _
      SQL = SQL & " and v.cpf = '" & CPF & "'"
   If Status <> "" Then _
      SQL = SQL & " and v.status = '" & Status & "'"
   TabVENDEDOR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub GERA_NUMR_REQ_DEV()
'On Error GoTo ERRO_TRATA

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_pedido = seq_pedido + 1 "

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   SQL = SQL & " AND cgc = '" & Trim(CNPJ_GERAL) & "'"

   CONECTA_RETAGUARDA.Execute SQL

   NUMR_REQ_DEV_N = 1

   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

   SQL = "select seq_pedido from EMPRESA "

   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEmpresa.EOF Then _
      If Not IsNull(TabEmpresa.Fields(0).Value) Then _
         NUMR_REQ_DEV_N = TabEmpresa.Fields(0).Value + 1
   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_NUMR_REQ_DEV"
End Sub

Public Sub GERA_NUMR_LOTE()
'On Error GoTo ERRO_TRATA

   NUMR_LOTE_N = MAX_ID("seq_lote", "EMPRESA", "", "", "", "")

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_lote = seq_lote + 1 "

   'SQL = SQL & " from EMPRESA "
   'SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   'SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_NUMR_LOTE"
End Sub

Public Sub GERA_NUMR_PEDIDO_COMPRA()
'On Error GoTo ERRO_TRATA

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_pedcompra = seq_pedcompra + 1 "

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   NUMR_COMPRA_N = 1

   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

   SQL = "select seq_pedcompra from EMPRESA "

   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEmpresa.EOF Then _
      If Not IsNull(TabEmpresa.Fields(0).Value) Then _
         NUMR_COMPRA_N = TabEmpresa.Fields(0).Value + 1
   If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_NUMR_PEDIDO_COMPRA"
End Sub

Public Sub GERA_CODIGO_PRODUTO()
RODA_PRODUTO:

   NUMR_PROD_N = 1

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select seq_codg_prod from EMPRESA "

   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      If Not IsNull(TabTemp.Fields(0).Value) Then _
         If IsNumeric(TabTemp.Fields(0).Value) Then _
            NUMR_PROD_N = 1 + TabTemp.Fields(0).Value

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_codg_prod = " & NUMR_PROD_N

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

   SQL = "select codg_produto from PRODUTO"
   SQL = SQL & " where codg_produto = '" & NUMR_PROD_N & "'"
   SQL = SQL & " and situacao <> 'C' "
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then _
      GoTo RODA_PRODUTO
End Sub

'Fun��es para Validar CPF e CNPJ
'Validar CGC
Public Function CALCULACGC(Numero As String) As String
   Dim i As Integer
   Dim prod As Integer
   Dim mult As Integer
   Dim Digito As Integer

   If Not IsNumeric(Numero) Then
      CALCULACGC = ""
      Exit Function
   End If

   mult = 2
   For i = Len(Numero) To 1 Step -1
     prod = prod + Val(Mid(Numero, i, 1)) * mult
     mult = IIf(mult = 9, 2, mult + 1)
   Next

   Digito = 11 - Int(prod Mod 11)
   Digito = IIf(Digito = 10 Or Digito = 11, 0, Digito)

   CALCULACGC = Trim(Str(Digito))
End Function

Public Function VALIDACGC(CGC As String) As Boolean
   If CALCULACGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
      VALIDACGC = False
      Exit Function
   End If
   If CALCULACGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
      VALIDACGC = False
      Exit Function
   End If
   VALIDACGC = True
End Function

Public Sub SP_PROCURA_ENDERE�ONFE(prop As String, tipo As String, CEP As String, IE_ID As Long)

   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   SQL = "select * from ENDERECO e "
   SQL = SQL & " inner join CEP c "
   SQL = SQL & " on e.cep = c.cep "
   SQL = SQL & " where e.PROP = '" & prop & "'"
   SQL = SQL & " and e.tipo in ('" & Trim(tipo) & "')"
   If Trim(CEP) <> "" Then _
      SQL = SQL & " and e.cep = '" & CEP & "'"
   If IE_ID > 0 Then _
      SQL = SQL & " and e.ie_id = " & IE_ID

   tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub SP_PROC_BANCO(Codg_Banco As String)
   SQL = "select * from BANCO "
   SQL = SQL & " where codg_banco = '" & Trim(Codg_Banco) & "'"
   TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub SP_GRAVA_CEP(CEP_A As String, CIDADE_A As String, UF_A As String, CODG_IBGE_N As Long)
   If TabCEP.State = 1 Then _
      TabCEP.Close

   If Trim(CEP_A) <> "" And Trim(CIDADE_A) <> "" And Trim(UF_A) <> "" Then
      SQL = "select * from CEP "
      SQL = SQL & " where cep = '" & Trim(CEP_A) & "'"
      SQL = SQL & " And cidade = '" & Trim(CIDADE_A) & "'"
      SQL = SQL & " And uf = '" & Trim(UF_A) & "'"
      TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCEP.EOF Then
         SQL = "insert into CEP "
            SQL = SQL & " (CEP,Cidade,UF,CODIGO_IBGE) "
         SQL = SQL & " values("
            SQL = SQL & "'" & Trim(CEP_A) & "'"
            SQL = SQL & ",'" & Trim(CIDADE_A) & "'"
            SQL = SQL & ",'" & Trim(UF_A) & "'"
            SQL = SQL & "," & Trim(CODG_IBGE_N)
         SQL = SQL & " )"
         Else
            SQL = "update CEP set"
               SQL = SQL & " CEP = '" & Trim(CEP_A) & "'"
               SQL = SQL & ", Cidade = '" & Trim(CIDADE_A) & "'"
               SQL = SQL & ", UF = '" & Trim(UF_A) & "'"
               SQL = SQL & ", CODIGO_IBGE = " & Trim(CODG_IBGE_N)
            SQL = SQL & " where cep = '" & Trim(CEP_A) & "'"
            SQL = SQL & " And cidade = '" & Trim(CIDADE_A) & "'"
            SQL = SQL & " And uf = '" & Trim(UF_A) & "'"
      End If

      CONECTA_RETAGUARDA.Execute SQL
   End If

   If TabCEP.State = 1 Then _
      TabCEP.Close
End Sub
'==============================CEP
Public Sub SP_PROCURA_CEP(CEP_A As String)
   If TabCEP.State = 1 Then _
      TabCEP.Close

   SQL = "select * from CEP"
   SQL = SQL & " where cep = '" & Trim(CEP_A) & "'"
   TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub SP_GRAVA_ENDERE�O(prop As String, _
                             CEP As String, _
                             Rua As String, _
                             Bairro As String, _
                             Complemento As String, _
                             tipo As String, _
                             IE_ID As Long, _
                             Numero As String)

   SQL = "delete ENDERECO "
   SQL = SQL & " where PROP = '" & prop & "'"
   SQL = SQL & " and tipo = '" & Trim(tipo) & "'"
   CONECTA_RETAGUARDA.Execute SQL

   SP_PROCURA_ENDERE�ONFE prop, tipo, CEP, IE_ID

   If Not tabEndereco.EOF Then
      SqL2 = "EXEC SP_UPDATE_END '" & prop & "','" & CEP & "','" & Rua & "','" & Bairro & "'," & Complemento & "','" & tipo & "'," & IE_ID & ",'" & Numero & "'"
      Else
         ENDERECO_ID_N = MAX_ID("ENDERECO_ID", "endereco", "", "", "", "")
         SqL2 = "EXEC SP_INSERT_END '" & prop & "','" & CEP & "','" & Rua & "','" & Bairro & "','" & Complemento & "','" & tipo & "'," & IE_ID & "," & ENDERECO_ID_N & ",'" & Numero & "'" & "," & PESSOA_ID_N
   End If
   CONECTA_RETAGUARDA.Execute SqL2

   If tabEndereco.State = 1 Then _
      tabEndereco.Close
End Sub

Public Sub SP_MATA_ENDERE�O(prop As String, _
                            tipo As String)
   SQL = "delete ENDERECO "
   SQL = SQL & " where PROP = '" & prop & "'"
   SQL = SQL & " and tipo = '" & Trim(tipo) & "'"
   CONECTA_RETAGUARDA.Execute SQL
End Sub

Public Sub SP_PROCURA_PRODUTO(Empresa As Long, _
                              Codg_Prod As String, _
                              Grupo As Long, _
                              Referencia As String, _
                              CGCCPF As String, _
                              Codg_Barra As String, _
                              Tipo_Prod As Integer)

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select p.*, d.*"
   SQL = SQL & " from PRODUTO p "
   SQL = SQL & " left join DESCR d "
   SQL = SQL & " on p.familiaproduto_id = d.codigo "
   SQL = SQL & " where p.empresa_id =  " & Empresa
   If Codg_Prod <> "" Then _
      SQL = SQL & " and p.codg_produto = '" & Codg_Prod & "'"
   If Grupo > 0 Then
      SQL = SQL & " and p.familiaproduto_id = " & Grupo
      SQL = SQL & " and d.tipo = 'G'"
   End If
   If Referencia <> "" Then _
      SQL = SQL & " and p.referencia = '" & Referencia & "'"
   If CGCCPF <> "" Then _
      SQL = SQL & " and p.cgccpf = '" & CGCCPF & "'"
   If Codg_Barra <> "" Then _
      SQL = SQL & " and p.CODG_BARRAS = '" & CODG_BARRAS & "'"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub CentralizaJanela2(Form As Form)
    Form.Top = (Screen.Height - Form.Height) / 2
    Form.Left = (Screen.Width - Form.Width) / 2
End Sub

Public Sub SP_PROCURA_FONE(prop As String, NUMR_FONE As Long)
   SqL2 = "EXEC SP_PROC_FONE '" & prop & "'"
   If TabFone.State = 1 Then TabFone.Close
   TabFone.Open SqL2, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub MOSTRA_MSG(Msg1 As String, Msg2 As Variant)
   frmINICIO.BARI.Panels.Clear
   frmINICIO.BARI.Panels.Add (1)
   frmINICIO.BARI.Panels(1).Text = Msg1
   frmINICIO.BARI.Panels(1).AutoSize = sbrContents

   If Trim(Msg2) <> "" Then
      frmINICIO.BARI.Panels.Add (2)
      frmINICIO.BARI.Panels(2).Text = Msg2
      frmINICIO.BARI.Panels(2).AutoSize = sbrContents
   End If

   frmINICIO.BARI.Refresh
End Sub

Public Function tpMOEDA(ByVal dado As Variant) As String
    tpMOEDA = Str(Trim$(IIf(IsNull(dado) Or Not IsNumeric(dado) Or Len(dado) = 0, 0, dado)))
End Function

Public Function tpTXT(ByVal dado As Variant, ByVal tamanho As Integer) As String
    tpTXT = Trim$(Left$(IIf(Len(dado) = 0 Or IsNull(dado), "", dado), tamanho))
End Function

Public Function ZerosEsquerda(ByVal Numero As Variant, ByVal q As Integer) As Variant
    If (Not IsNull(Numero)) Or (Len(Numero) > 0) Then ZerosEsquerda = Format$(Numero, Left$("0000000000", q))
End Function

Public Function numeros(ByVal texto As String) As String
    Dim i As Long
    numeros = ""
    For i = 1 To Len(texto)
        If IsNumeric(Mid(texto, i, 1)) Then numeros = numeros & Mid(texto, i, 1)
    Next i
End Function

Public Function tpDT(ByVal dado As Variant) As String
    tpDT = mda(IIf(Len(dado) = 0 Or Not IsDate(dado) Or IsNull(dado), "31/12/1601", dado))
End Function

Public Function tpINT(ByVal dado As Variant) As Long
    tpINT = IIf(Len(dado) = 0 Or Not IsNumeric(dado) Or IsNull(dado), 0, dado)
End Function

'##ModelId=417E3D74013D
Public Function MaiusculasLetrasNumeros(ByVal Enter As Integer) As Integer
    MaiusculasLetrasNumeros = Asc(UCase$(Chr(Enter)))
End Function

'##ModelId=417E3D74014D
Public Function tpLNG(ByVal dado As Variant) As Long
    tpLNG = CLng(IIf(Len(dado) = 0 Or Not IsNumeric(dado) Or IsNull(dado), 0, dado))
End Function

'##ModelId=417E3D740052
Public Function PreencheObjetos(adoControl As Object, strsql As String, grade As Variant) As String
On Error Resume Next

    adoControl.ConnectionString = strConexao
    adoControl.UserName = uid
    adoControl.Password = pwd
    adoControl.RecordSource = strsql
    adoControl.CommandTimeout = 200
    grade.Refresh
    adoControl.Refresh
    
'    Exit Function
'ERRO_TRATA:
'    MsgBox Err.Description
End Function

Public Function Enter(X As Integer)
    If X = vbKeyReturn Then
        X = 0
        SendKeys "{tab}"
    End If
End Function

Public Sub calcula_dias_uteis(datai As String, dataf As String)
    Dim data_aux As Date
    Dim data_atual As Date
    Dim ultimo_dia_mes As String
    Dim feriado As Integer
    Dim Rs As New ADODB.Recordset
    
   If Month(datai) <> Month(dataf) Then
      dataf = UltimoDiaMes(Month(datai), Year(dataf))
   End If

   dias_uteis = 0
   dias_realizados = 0
   ' marcado
   Rs.Open "SELECT getdate() as data", CONECTA_RETAGUARDA, , adCmdText
   If (DateValue(Rs("data")) <= DateValue(dataf)) Then
      data_atual = Rs("data")
   Else
      data_atual = dataf
   End If
   Rs.Close

   ultimo_dia_mes = "01/" & Format(DateAdd("m", 1, data_atual), "mm/yyyy")

   Rs.Open "SELECT data FROM feriado WHERE data >= '" & DMA(datai) & "' and data < '" & DMA(ultimo_dia_mes) & "'", CONECTA_RETAGUARDA, , adCmdText
   If (Not Rs.EOF) Then
      Rs.MoveFirst
      While (Not Rs.EOF)
         If (data_atual >= Rs("data")) Then
            dias_realizados = dias_realizados - 1
         End If
         feriado = feriado + 1
         dias_uteis = dias_uteis - 1
         Rs.MoveNext
      Wend
   End If
   Rs.Close

   data_aux = datai

   If (DateValue(data_aux) < DateValue(ultimo_dia_mes)) Then
      While (DateValue(data_aux) < DateValue(ultimo_dia_mes))
         If (Weekday(data_aux) <> 1 And Weekday(data_aux) <> 7) Then
            dias_uteis = dias_uteis + 1
            If (DateValue(Date) >= DateValue(data_aux)) Then
               dias_realizados = dias_realizados + 1
            End If
         End If
         data_aux = DateAdd("d", 1, data_aux)
      Wend
   Else
      MsgBox "Per�odo incorreto...", vbCritical, "Aten��o !!."
   End If
End Sub

Public Function ValidaInscricaoEstadual(Insc As String, UF As String) As Integer
'On Error GoTo ERRO_TRATA

   If UCase(Insc) = "ISENTA" Then _
      Insc = "ISENTO"

   strRetorno = Insc
   Insc = Trim(Insc)
   Insc = Replace(Insc, ".", "")
   Insc = Replace(Insc, ",", "")
   Insc = Replace(Insc, "-", "")
   Insc = Replace(Insc, "/", "")

   ValidaInscricaoEstadual = Inscricao(Trim(Insc), Trim(UF))

   If ValidaInscricaoEstadual = 1 Then
       MsgBox "Inscricao Inv�lida para a Uf Informada . " & strRetorno, vbExclamation, "SHFSYS"
   ElseIf ValidaInscricaoEstadual = 2 Then
       MsgBox "Cliente com parametros de inscricao estadual inv�lidos ." & strRetorno, vbExclamation, "SHFSYS"
   End If

Exit Function
ERRO_TRATA:
   MsgBox Err.Description
End Function

'FUN��O QUE RETORNA O NUMERO DE DIAS NO M�S
Function nDiaMes(dData) As Integer
    nDiaMes = 32 - Day(CDate("01/" & Month(dData) & "/" & Year(dData)) + 31)
End Function

Public Function TiraAcento(ByVal texto As String) As String
'On Error GoTo ERRO_TRATA

   Dim C As String
   Dim k As Integer
   C = ""
   For k = 1 To Len(texto)
      If Trim(UCase$(Mid(texto, k, 1))) = "" Then
         C = C & " "
         Else
            Select Case Trim(UCase$(Mid(texto, k, 1)))
               Case "�", "�", "�", "�", "�", "�"
                  C = C & "A"
               Case "�", "�", "�", "�"
                  C = C & "E"
               Case "�", "�", "�"
                  C = C & "I"
               Case "�", "�", "�", "�", "�"
                  C = C & "O"
               Case "�", "�", "�", "�"
                  C = C & "U"
               Case "�"
                  C = C & "C"
               Case "�"
                  C = C & "N"
               Case "�", "�"
                  C = C & "Y"
               Case "�", "`", "'"
                  C = C & " "
               Case ""
                  C = C & " "
               Case Else
                  C = C & Mid(texto, k, 1)
            End Select
      End If
   Next k
   TiraAcento = UCase$(Trim$(C))

   Exit Function

ERRO_TRATA:
   MsgBox Err.Description
End Function

Public Function CaracteresValidos(ByVal texto As String) As String
    Dim C As String
    Dim k As Integer
    C = ""
    For k = 1 To Len(texto)
        If Asc(Mid(texto, k, 1)) >= 32 And Asc(Mid(texto, k, 1)) <= 127 Then
            If (Asc(Mid(texto, k, 1)) >= 32 And Asc(Mid(texto, k, 1)) <= 32) Or _
               (Asc(Mid(texto, k, 1)) >= 40 And Asc(Mid(texto, k, 1)) <= 41) Or _
               (Asc(Mid(texto, k, 1)) >= 48 And Asc(Mid(texto, k, 1)) <= 57) Or _
               (Asc(Mid(texto, k, 1)) >= 65 And Asc(Mid(texto, k, 1)) <= 90) Or _
               (Asc(Mid(texto, k, 1)) >= 44 And Asc(Mid(texto, k, 1)) <= 46) Then
                C = C & Mid(texto, k, 1)
            End If
        End If
    Next k
    CaracteresValidos = Trim$(C)
End Function
'##ModelId=417E3D750043
Public Sub ConfiguraGrid(Grid As Variant)
   Grid.Override.AllowColSizing = ssAllowColSizingFree
   Grid.Override.HeaderClickAction = ssHeaderClickActionSortMulti

   'ordena
   Grid.ViewStyleBand = ssViewStyleBandVertical
   Grid.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
   Grid.Override.FetchRows = ssFetchRowsPreloadWithParent
   Grid.Override.HeaderClickAction = ssHeaderClickActionSortMulti

   Grid.AlphaBlendEnabled = True
   Grid.Override.CellClickAction = ssClickActionRowSelect
   Grid.Override.SelectedCellAppearance.BackColorAlpha = ssAlphaUseAlphaLevel

   Grid.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
End Sub

Public Function MontaSQLRelatorio(lngCodigoRelatorio As Long, Botao As Object) As String
'On Error GoTo ERRO_TRATA

   Dim rsBotao As New ADODB.Recordset
   
   Dim strFiltroRelatorio As String
   
   rsBotao.Open "SELECT * from RelatorioFiltro WHERE Codigo = " & lngCodigoRelatorio & " AND CodigoUsuario = " & intUsuario, CONECTA_RETAGUARDA, , adCmdText
   If Not rsBotao.EOF Then
       strFiltroRelatorio = Trim(rsBotao!SelectionFormula & "")
   Else
       MontaSQLRelatorio = ""
       'MsgBox "Consulta inexistente...", 48, "Aten��o..."
       rsBotao.Close
       GoTo Sai
   End If
   rsBotao.Close
   strFiltroRelatorio = Replace(strFiltroRelatorio, "'", "'")
   
Sai:
   If strFiltroRelatorio = "" Then
       Botao.Picture = LoadPicture(App.Path & "\Imagem\FiltroVerde.gif")
   Else
       Botao.Picture = LoadPicture(App.Path & "\Imagem\FiltroVermelho.gif")
   End If
   MontaSQLRelatorio = strFiltroRelatorio
    
   Exit Function

ERRO_TRATA:
    'ControleErros Err.Number, Err.Description, Err.Source, "Monta bot�a verde e vermelho..."
End Function

Public Function Relatorio(strFormulario As String, strRelatorio As String, strNomeRelatorio As String, strFiltro As String, Optional strParametro1 As String, Optional strParametro2 As String, Optional strParametro3 As String, Optional strParametro4 As String, Optional strParametro5 As String, Optional strParametro6 As String, Optional strParametro7 As String, Optional strParametro8 As String, Optional strParametro9 As String, Optional strParametro10 As String, Optional strParametro11 As String, Optional strParametro12 As String, Optional strParametro13 As String, Optional strParametro14 As String, Optional strParametro15 As String, Optional strParametro16 As String, Optional strParametro17 As String, Optional strParametro18 As String, Optional strParametro19 As String, Optional strParametro20 As String, Optional strParametro21 As String, Optional strParametro22 As String, Optional strParametro23 As String)
'On Error GoTo ERRO_TRATA

   Dim LS_Report As String
   Dim gsNomeRel As String
   Dim sParametro As String
   Dim sSelectionFormula As String
   Dim strImpressora As String
   
   sSelectionFormula = CRITERIO
   
   'If strImpressora = "" Then
   '    strRetorno1 = strNomeRelatorio
   '    frmImpressora.Show 1
   'End If
   
   'If strRetorno = "" Then Exit Function
   'strImpressora = strRetorno
   
   If UCase(Right(strRelatorio, 3)) = "RPT" Then
       LS_Report = PATH_REL & strRelatorio
     Else: LS_Report = App.Path & "\Relatorio\" & strRelatorio & ".rpt"
   End If
   
   gsNomeRel = strNomeRelatorio

   Set crxReport = crxApplication.OpenReport(LS_Report)
   crxReport.DiscardSavedData
     
   'If frmImpressora.chkConexao.Value = 1 Then
   '    GS_Define_Usuario_Senha
   'End If

   crxReport.RecordSelectionFormula = sSelectionFormula
   crxReport.GroupSelectionFormula = SelectionFormulaGrupo

   If strParametro1 <> "" Then LS_Envia_Formula Left(strParametro1, InStr(1, strParametro1, "=") - 1), Mid(strParametro1, InStr(1, strParametro1, "=") + 1, Len(strParametro1))
   If strParametro2 <> "" Then LS_Envia_Formula Left(strParametro2, InStr(1, strParametro2, "=") - 1), Mid(strParametro2, InStr(1, strParametro2, "=") + 1, Len(strParametro2))
   If strParametro3 <> "" Then LS_Envia_Formula Left(strParametro3, InStr(1, strParametro3, "=") - 1), Mid(strParametro3, InStr(1, strParametro3, "=") + 1, Len(strParametro3))
   If strParametro4 <> "" Then LS_Envia_Formula Left(strParametro4, InStr(1, strParametro4, "=") - 1), Mid(strParametro4, InStr(1, strParametro4, "=") + 1, Len(strParametro4))
   If strParametro5 <> "" Then LS_Envia_Formula Left(strParametro5, InStr(1, strParametro5, "=") - 1), Mid(strParametro5, InStr(1, strParametro5, "=") + 1, Len(strParametro5))
   If strParametro6 <> "" Then LS_Envia_Formula Left(strParametro6, InStr(1, strParametro6, "=") - 1), Mid(strParametro6, InStr(1, strParametro6, "=") + 1, Len(strParametro6))
   If strParametro7 <> "" Then LS_Envia_Formula Left(strParametro7, InStr(1, strParametro7, "=") - 1), Mid(strParametro7, InStr(1, strParametro7, "=") + 1, Len(strParametro7))
   If strParametro8 <> "" Then LS_Envia_Formula Left(strParametro8, InStr(1, strParametro8, "=") - 1), Mid(strParametro8, InStr(1, strParametro8, "=") + 1, Len(strParametro8))
   If strParametro9 <> "" Then LS_Envia_Formula Left(strParametro9, InStr(1, strParametro9, "=") - 1), Mid(strParametro9, InStr(1, strParametro9, "=") + 1, Len(strParametro9))
   If strParametro10 <> "" Then LS_Envia_Formula Left(strParametro10, InStr(1, strParametro10, "=") - 1), Mid(strParametro10, InStr(1, strParametro10, "=") + 1, Len(strParametro10))

   If strParametro11 <> "" Then LS_Envia_Formula Left(strParametro11, InStr(1, strParametro11, "=") - 1), Mid(strParametro11, InStr(1, strParametro11, "=") + 1, Len(strParametro11))
    If strParametro12 <> "" Then LS_Envia_Formula Left(strParametro12, InStr(1, strParametro12, "=") - 1), Mid(strParametro12, InStr(1, strParametro12, "=") + 1, Len(strParametro12))
    If strParametro13 <> "" Then LS_Envia_Formula Left(strParametro13, InStr(1, strParametro13, "=") - 1), Mid(strParametro13, InStr(1, strParametro13, "=") + 1, Len(strParametro13))
    If strParametro14 <> "" Then LS_Envia_Formula Left(strParametro14, InStr(1, strParametro14, "=") - 1), Mid(strParametro14, InStr(1, strParametro14, "=") + 1, Len(strParametro14))
    If strParametro15 <> "" Then LS_Envia_Formula Left(strParametro15, InStr(1, strParametro15, "=") - 1), Mid(strParametro15, InStr(1, strParametro15, "=") + 1, Len(strParametro15))
    If strParametro16 <> "" Then LS_Envia_Formula Left(strParametro16, InStr(1, strParametro16, "=") - 1), Mid(strParametro16, InStr(1, strParametro16, "=") + 1, Len(strParametro16))
    If strParametro17 <> "" Then LS_Envia_Formula Left(strParametro17, InStr(1, strParametro17, "=") - 1), Mid(strParametro17, InStr(1, strParametro17, "=") + 1, Len(strParametro17))
    If strParametro18 <> "" Then LS_Envia_Formula Left(strParametro18, InStr(1, strParametro18, "=") - 1), Mid(strParametro18, InStr(1, strParametro18, "=") + 1, Len(strParametro18))
    If strParametro19 <> "" Then LS_Envia_Formula Left(strParametro19, InStr(1, strParametro19, "=") - 1), Mid(strParametro19, InStr(1, strParametro19, "=") + 1, Len(strParametro19))
    If strParametro20 <> "" Then LS_Envia_Formula Left(strParametro20, InStr(1, strParametro20, "=") - 1), Mid(strParametro20, InStr(1, strParametro20, "=") + 1, Len(strParametro20))

    If strParametro21 <> "" Then LS_Envia_Formula Left(strParametro21, InStr(1, strParametro21, "=") - 1), Mid(strParametro21, InStr(1, strParametro21, "=") + 1, Len(strParametro21))
    If strParametro22 <> "" Then LS_Envia_Formula Left(strParametro22, InStr(1, strParametro22, "=") - 1), Mid(strParametro22, InStr(1, strParametro22, "=") + 1, Len(strParametro22))
    If strParametro23 <> "" Then LS_Envia_Formula Left(strParametro23, InStr(1, strParametro23, "=") - 1), Mid(strParametro23, InStr(1, strParametro23, "=") + 1, Len(strParametro23))
   
    LogRelatorio strFormulario, LS_Report, Left(gsNomeRel, 200), Left(TrocaApostrofeSharpe(strParametro1), 200), Left(TrocaApostrofeSharpe(sSelectionFormula), 200)
    If strImpressora = "VIDEO" Then
        Dim frmRel As New frmRELATORIO10
        
        'frmRel.crvRelatorio.ReportSource = crxReport
        frmRel.Caption = gsNomeRel
        'frmRel.crvRelatorio.ViewReport
        'frmRel.crvRelatorio.Zoom 105
        'frmRel.crvRelatorio.DisplayGroupTree = True
        If UCase(Left(strRelatorio, 6)) = "EMITEN" Or UCase(Left(strRelatorio, 6)) = "TABELA" Or UCase(Left(strRelatorio, 14)) = "ESTOQUECRITICA" Or UCase(Left(strRelatorio, 18)) = "FATURAMENTOESPELHO" Or UCase(Left(strRelatorio, 18)) = "FINANCEIROREMESSA" Or UCase(Left(strRelatorio, 5)) = "NOTAS" Or UCase(Left(strRelatorio, 7)) = "BALANCO" Or strNomeRelatorio = "Boletos" Or UCase(Left(strRelatorio, 15)) = "FORMULARIOCOTAC" Or UCase(Left(strRelatorio, 15)) = "COMPROVANTEBAIX" Or UCase(Left(strRelatorio, 6)) = "RECIBO" Or UCase(Left(strRelatorio, 3)) = "BEM" Or UCase(Left(strRelatorio, 11)) = "FATURAMENTO" Or UCase(Left(strRelatorio, 6)) = "FINANC" Then
            'Msg ""
            frmRel.Show 1
        Else
            'tava dando pau aq, a� coloquei o 1 na frente = bica
            'Msg ""
            frmRel.Show 1
        End If
    ElseIf strImpressora = "ARQUIVO" Then
        'crxReport.ExportOptions.DiskFileName = "c:\" & Left(strRelatorio, Len(strRelatorio) - 4) & ".doc"
        'crxReport.ExportOptions.DestinationType = crEDTDiskFile 'or use crEDTDiskFile
        'crxReport.ExportOptions.FormatType = crEFTWordForWindows 'format changes for .rft, word, PDF, etc
        crxReport.Export True
        'crxReport.SaveAs "d:\ivrs\crystalreports\subjects15.rpt", cr80FileFormat
        'crxReport.SaveAs "C:\" & Left(strRelatorio, Len(strRelatorio) - 4) & ".XLS", crDefaultFileFormat
    Else
        crxReport.PrinterSetup frmINICIO.hWnd
        crxReport.PrintOut False, 1, True
        
        If MsgBox("Confirma Impress�o?", vbQuestion + vbYesNo, "Gera Faturamento") = vbNo Then 'Wanderson - 25-02-2010
            Exit Function
        End If
        
'        If crxReport.DriverName <> "" Then
'            Printer.Orientation = crxReport.PaperOrientation
'            crxReport.DisplayProgressDialog = True
'            crxReport.SelectPrinter crxReport.DriverName, crxReport.PrinterName, crxReport.PortName
'            crxReport.PaperOrientation = Printer.Orientation
'            crxReport.PrintOut True, 1
'        End If
    End If

    Exit Function
ERRO_TRATA:
    If Err.Number = "-2147189423" Then
        MsgBox "Aten��o. Selecione uma impressora padrao."
        Exit Function
    ElseIf Err.Number = "-2147206461" Then
        MsgBox "Aten��o. Relat�rio nao localizado Caminho: " & LS_Report
        Exit Function
    End If
'    ControleErros Err.Number, Err.Description, Err.Source, "GeraRelatorio"
End Function

'##ModelId=417E3D7402F4
Public Sub LS_Envia_Formula(LS_Formula1 As String, LS_Parametro1 As String)
    'configurar f�rmula
    Set OG_Formula_Field = crxReport.FormulaFields
    For Each OG_Formula_Field In OG_Formula_Field
        If Trim(UCase(OG_Formula_Field.Name)) = Trim(UCase(LS_Formula1)) Then
            OG_Formula_Field.Text = LS_Parametro1
            Exit Sub
        End If
    Next
End Sub

Public Sub LS_Envia_FormulaSubReport(LS_Formula1 As String, LS_Parametro1 As String)
    'configurar f�rmula
    Set OG_Formula_Field = crxSubReport.FormulaFields
    For Each OG_Formula_Field In OG_Formula_Field
        If Trim(UCase(OG_Formula_Field.Name)) = Trim(UCase(LS_Formula1)) Then
            OG_Formula_Field.Text = LS_Parametro1
            Exit Sub
        End If
    Next
End Sub

Public Sub LS_Envia_Parametro(LS_Formula1 As String, LS_Parametro1 As String)
'On Error GoTo ERRO_TRATA

   For i = 1 To crxReport.ParameterFields.Count
      'inteiro
      If crxReport.ParameterFields.Item(i).Name = LS_Formula1 And crxReport.ParameterFields.Item(i).ValueType = 7 Then
         crxReport.ParameterFields.Item(i).AddCurrentValue CDbl(LS_Parametro1)
         Exit Sub
      End If
      'valor
      If crxReport.ParameterFields.Item(i).Name = LS_Formula1 And crxReport.ParameterFields.Item(i).ValueType = 8 Then
         crxReport.ParameterFields.Item(i).AddCurrentValue CDbl(LS_Parametro1)
         Exit Sub
      End If
      'texto
      If crxReport.ParameterFields.Item(i).Name = LS_Formula1 And crxReport.ParameterFields.Item(i).ValueType = 12 Then
         crxReport.ParameterFields.Item(i).AddCurrentValue (LS_Parametro1)
         Exit Sub
      End If
      'date
      If crxReport.ParameterFields.Item(i).Name = LS_Formula1 And (crxReport.ParameterFields.Item(i).ValueType = 16 Or crxReport.ParameterFields.Item(i).ValueType = 10) Then
         crxReport.ParameterFields.Item(i).AddCurrentValue (DateValue(LS_Parametro1))
         Exit Sub
      End If
   Next i
   
   Exit Sub
ERRO_TRATA:
    MsgBox Err.Description
End Sub
'##ModelId=417E3D740295
Public Sub ExportaExcel(adoConsulta As ADODB.Recordset, NomeArquivo As String)
'On Error GoTo ERRO_TRATA

   'verifica se existe o diretorio
   'If Dir(App.Path & "\Arquivos\*.xls") = "" Then _
      MkDir App.Path & "\Arquivos"

   Dim oExcel  As Object
   Dim oBook   As Object
   Dim oSheet  As Object
   Dim n       As Integer

   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)

   'Transfer the field names to Row 1 of the worksheet:
   'Note: CopyFromRecordset copies only the data and not the field
   '      names, so you can transfer the fieldnames by traversing the
   '      fields collection.
   For n = 1 To adoConsulta.Fields.Count
      oSheet.Cells(1, n).Value = adoConsulta.Fields(n - 1).Name
   Next

   'Transfer the data to Excel.
   oSheet.Range("A2").CopyFromRecordset adoConsulta

   'Save the workbook and quit Excel.
   oBook.SaveAs NomeArquivo
   oExcel.Visible = True

   Set oSheet = Nothing
   Set oBook = Nothing
   Set oExcel = Nothing

   MsgBox "Exportado com sucesso... Nome Arquivo: " & NomeArquivo

   oExcel.Quit

Exit Sub
ERRO_TRATA:
   If Err.Number = 75 Then _
       Resume Next
   If Err.Number <> 0 And Err.Number <> 91 Then _
      MsgBox "Excel Ja Aberto, favor fechar e abrir de novo." & Chr(13) & Chr(10) & Err.Description, 48, "Exporta Excel"
End Sub

Public Sub WordExportacao()
'On Error GoTo CheckWord
On Error Resume Next

Dim wd As New Word.Application
Dim doc As Word.Document
Dim tb As Word.Table
Dim f As ADODB.Field
Dim k As Integer
Dim i As Integer

   If cnn.State = 1 Then cnn.Close

   cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtArquivo.Text
   Rs.CursorLocation = adUseClient
   Rs.Open "select * from " & cboTabela.Text, cnn, 2, 3
   pb.Max = Rs.RecordCount
   pb.Value = 0
   Screen.MousePointer = vbHourglass
   Set doc = wd.Documents.Add
   Set tb = doc.Tables.Add(wd.Selection.Range, Rs.RecordCount + 1, Rs.Fields.Count)

    k = 1
    For Each f In Rs.Fields
           tb.Cell(1, k).Range.Font.Bold = True
           tb.Cell(1, k).Range.Text = f.Name
           k = k + 1
   Next
   Rs.MoveFirst

   For i = 1 To Rs.RecordCount
      pb.Value = i
      k = 1
       For Each f In Rs.Fields
           tb.Cell(i + 1, k).Range.Text = Rs.Fields(f.Name)
            k = k + 1
       Next
       Rs.MoveNext
   Next

   Screen.MousePointer = vbNormal
   MsgBox "Convers�o realizada com sucesso"
   pb.Value = 0

   doc.SaveAs txtDestino.Text
   Set doc = Nothing
   Set w = Nothing
   Rs.Close
   Set Rs = Nothing
   cnn.Close
   Set cnn = Nothing
   'Exit Sub
   'CheckWord:
   'MsgBox Err.Description
End Sub
'##ModelId=417E3D74017D
'Public Function tpTXT(ByVal dado As Variant, ByVal tamanho As Integer) As String
'    tpTXT = Trim$(Left$(IIf(Len(dado) = 0 Or IsNull(dado), "", dado), tamanho))
'End Function

'##ModelId=417E3D74019B
'Public Function tpmoeda(ByVal dado As Variant) As String
'    moeda = Str(Trim$(IIf(IsNull(dado) Or Not IsNumeric(dado) Or Len(dado) = 0, 0, dado)))
'End Function

'##ModelId=417E3D7401AA
Public Function tpHS(ByVal dado As Variant) As String
    tpHS = hhmmss(IIf(Len(dado) = 0 Or IsNull(dado), 0, dado))
End Function

'##ModelId=417E3D7401BA
Public Function hhmmss(ByRef hora As Variant) As String
    hhmmss = Format$(hora, "HH:MM:SS")
End Function

'##ModelId=417E3D7401CA
Public Function hhmm(ByRef hora As Variant) As String
    hhmm = Format$(hora, "HH:MM")
End Function

'##ModelId=417E3D7401D9
Public Function tpCEP(ByVal dado As Variant) As String
    If Len(Trim(dado)) = 0 Then
        tpCEP = "  .   -   "
    ElseIf IsNull(dado) Then
        tpCEP = "  .   -   "
    Else
        dado = numeros(dado)
        tpCEP = Left(Format(dado, "&&.&&&-&&&"), 10)
    End If
End Function

Public Function tpFone(ByVal dado As Variant) As String
    If Len(Trim(dado)) = 0 Then
        tpFone = "(    )     -    "
    ElseIf IsNull(dado) Then
        tpFone = "(    )     -    "
    Else
        dado = numeros(dado)
        If Left(dado, strFormatacao2Digitos) = "62" Then
            tpFone = "(0XX62) " & Left(Format(Right(dado, Len(dado) - 2), "&&&&-&&&&"), 16)
        ElseIf Left(dado, 3) = "062" Then
            tpFone = "(0XX62) " & Left(Format(Right(dado, Len(dado) - 3), "&&&&-&&&&"), 16)
        Else
            tpFone = Left(Format(dado, "(&&&&) &&&&-&&&&"), 16)
        End If
        
    End If
End Function

'##ModelId=417E3D7401E9
Public Function LimpaCampos(ByRef Formulario As Form)
    On Error Resume Next
    Dim i As Long
    For i = 1 To Formulario.Count
        If Formulario.Controls(i).Tag <> "N" Then
            Formulario.Controls(i).Text = ""
        End If
    Next i
End Function

'##ModelId=417E3D7401F9
Public Function HabilitaCampos(Formulario As Form, Valor As Boolean, Optional NHabilita As Boolean)
    On Error Resume Next
    Dim i As Long
    For i = 1 To Formulario.Count
        If Formulario.Controls(i).Tag <> "N" Then
            Formulario.Controls(i).Enabled = Not Valor
        ElseIf Formulario.Controls(i).Tag = "N" And NHabilita = True Then
            Formulario.Controls(i).Enabled = Valor
        End If
    Next i
End Function

'##ModelId=417E3D74020A
Public Sub CentraNaTela(f As Form)
    With frmINICIO
        If f.WindowState = vbNormal Then              'se o form n�o est� minimizado
            If TypeOf f Is MDIForm Then                 'se for o MDI (principal)
                f.Top = (Screen.Height - f.Height) / 2    'centra na tela
                f.Left = (Screen.Width - f.Width) / 2
            Else
                If f.MDIChild = True Then                 'se for "filho" do principal
                    f.Top = (.ScaleHeight - f.Height) / 2
                    f.Left = (.ScaleWidth - f.Width) / 2    'calcula coordenadas do canto esquerdo
                End If
            End If
        End If
    End With
End Sub

Public Function DMA(ByRef Data As String, Optional Tipo_I_inicial_F_Final As String = "N") As String
    If Tipo_I_inicial_F_Final = "N" Then
        DMA = Format$(IIf(Not IsDate(Data), Date, Data), "dd/MM/yyyy") & strData
    ElseIf UCase(Tipo_I_inicial_F_Final) = "I" Then
        DMA = Format$(IIf(Not IsDate(Data), Date, Data), "dd/MM/yyyy 00:00:00") & strData
    ElseIf UCase(Tipo_I_inicial_F_Final) = "F" Then
        DMA = Format$(IIf(Not IsDate(Data), Date, Data), "dd/MM/yyyy 23:59:59") & strData
    End If
End Function

Public Function mda(ByRef Data As String, Optional Tipo_I_inicial_F_Final As String = "N") As String
    ' Fun��o      : Serve para formatar e retornar o mes dia e ano em formato Americano, e se passar o parametro retorna a hora
    ' Autor       :
    ' Altera��es  :

    If Tipo_I_inicial_F_Final = "N" Then 'retorna data formatada sem horas
        mda = Format$(IIf(Not IsDate(Data), Date, Data), "MM/dd/yyyy")
    ElseIf Tipo_I_inicial_F_Final = "I" Then 'retorna data formatada com a hora inicial do dia
        mda = Format$(IIf(Not IsDate(Data), Date, Data), "MM/dd/yyyy 00:00:00")
    ElseIf Tipo_I_inicial_F_Final = "F" Then 'retorna data formatada com a ultima hora dodia
        mda = Format$(IIf(Not IsDate(Data), Date, Data), "MM/dd/yyyy 23:59:59")
    ElseIf Tipo_I_inicial_F_Final = "H" Then  'retorna data formatada com a hora atual
        mda = Format$(IIf(Not IsDate(Data), Date, Data), "MM/dd/yyyy ") & Format(Time, "hh:MM:ss")
    End If
End Function

Public Function mdaI(ByRef Data As String) As String
    mdaI = strData & Format$(IIf(Not IsDate(Data), Date, Data), "MM/dd/yyyy 00:00:00") & strData
End Function

Public Function mdaF(ByRef Data As String) As String
    mdaF = strData & Format$(IIf(Not IsDate(Data), Date, Data), "MM/dd/yyyy 23:59:59") & strData
End Function

Public Function MOEDA_FORMAT(ByRef dado As String, Optional CasasDecimais As Integer = 2) As String
    ' Fun��o      : Converte um valor em Moeda
    ' Autor       :
    ' Altera��es  :
    If Not IsNumeric(dado) Then dado = 0
        
    If CasasDecimais = 0 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0")
    ElseIf CasasDecimais = 1 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.0")
    ElseIf CasasDecimais = 2 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "Standard")
    ElseIf CasasDecimais = 3 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.000")
    ElseIf CasasDecimais = 4 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.0000")
    ElseIf CasasDecimais = 5 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.00000")
    ElseIf CasasDecimais = 6 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.000000")
    ElseIf CasasDecimais = 7 Then
        MOEDA = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.0000000")
    End If
End Function

Public Function UltimoDiaMes(Mes As Integer, Ano As Integer) As Date
   Dim DataFinal As Date
   Dim Meses As Integer

   Ano = IIf(Mes <> 12, Ano, Ano + 1)
   Meses = IIf(Mes <> 12, Mes + 1, 1)
   DataFinal = "01/" & Format(Meses, "00") & "/" & Ano
   UltimoDiaMes = DataFinal - 1
End Function

'##ModelId=417E3D7402B5
Public Function TrocaApostrofeSharpe(Valor As String) As String
   For i = 1 To Len(Valor)
      If Mid(Valor, i, 1) <> "'" Then
         TrocaApostrofeSharpe = TrocaApostrofeSharpe & Mid(Valor, i, 1)
         Else: TrocaApostrofeSharpe = TrocaApostrofeSharpe & "'"
      End If
      DoEvents
   Next i
End Function

Public Sub LogRelatorio(Formulario As String, TipoRelatorio As String, NomeRpt As String, Periodo As String, SelectionFormula As String)
    Dim varsql As String
    varsql = "INSERT INTO LogRelatorio (CodigoEmpresa, DataHora, CodigoUsuario, Formulario, TipoRelatorio, NomeRpt, Periodo, SelectionFormula)"
    varsql = varsql & " VALUES ("
    varsql = varsql & intEmpresa
    varsql = varsql & ",getdate()" 'MARCADO PARA NAO MUDAR
    varsql = varsql & " , " & intUsuario
    varsql = varsql & " , '" & Formulario & "'"
    varsql = varsql & " , '" & TipoRelatorio & "'"
    varsql = varsql & " , '" & NomeRpt & "'"
    varsql = varsql & " , '" & Periodo & "'"
    varsql = varsql & " , '" & SelectionFormula & "'"
    varsql = varsql & ")"
    CONECTA_RETAGUARDA.Execute varsql
End Sub

'##ModelId=417E3D74010E
Public Function ValidaCPF(NUM_CPF As String) As Boolean
    On Error Resume Next
    '//------------------------------------
    '// Fun��o que testa se o CPF � v�lido
    '//------------------------------------

    Dim aux As String
    Dim CPF As String
    Dim Calculo As Double
    Dim resto As Integer
    Dim digitos As String

    If IsNumeric(NUM_CPF) Then
        NUM_CPF = Format$(NUM_CPF, "&&&.&&&.&&&-&&")
        If Len(NUM_CPF) <> 14 Then
            ValidaCPF = False
            Exit Function
        End If
    End If

    'Selecionando apenas os n�meros da stringue NUM_CPF
    For i = 1 To Len(Trim$(NUM_CPF))
        If Mid$(NUM_CPF, i, 1) <> "." And Mid$(NUM_CPF, i, 1) <> "-" Then
            CPF = CPF & Mid$(NUM_CPF, i, 1)
        End If
    Next i

    'Se o tamanho do CPF passado for < 11 ent�o de imediato j� � falso
    If Len(CPF) <> 11 Then
        ValidaCPF = False
        Exit Function
    End If

    aux = Right(CPF, 2) 'recebe os dois digitos para ser comparado

    CPF = Left(CPF, 9)

    'C�lculo do primeiro digito
    Calculo = 0
    Calculo = (Mid$(CPF, 9, 1) * 9) + (Mid$(CPF, 8, 1) * 8) + (Mid$(CPF, 7, 1) * 7) + (Mid$(CPF, 6, 1) * 6) + (Mid$(CPF, 5, 1) * 5) + (Mid$(CPF, 4, 1) * 4) + (Mid$(CPF, 3, 1) * 3) + (Mid$(CPF, 2, 1) * 2) + (Mid$(CPF, 1, 1) * 1)
    resto = (Calculo Mod 11)

    If resto = 10 Then
        resto = 0
    End If

    digitos = resto

    'C�lculo do segundo digito
    Calculo = 0
    CPF = CPF & resto
    Calculo = 0
    Calculo = (Mid$(CPF, 10, 1) * 9) + (Mid$(CPF, 9, 1) * 8) + (Mid$(CPF, 8, 1) * 7) + (Mid$(CPF, 7, 1) * 6) + (Mid$(CPF, 6, 1) * 5) + (Mid$(CPF, 5, 1) * 4) + (Mid$(CPF, 4, 1) * 3) + (Mid$(CPF, 3, 1) * 2) + (Mid$(CPF, 2, 1) * 1) + (Mid$(CPF, 1, 1) * 0)
    resto = (Calculo Mod 11)
    If resto = 10 Then
        resto = 0
    End If
    digitos = digitos & resto
    
    strRetorno = digitos
    If aux <> digitos Then
        ValidaCPF = False
        strRetorno = digitos
       Else: ValidaCPF = True
    End If
End Function

Public Sub CarregarRelatorio()
On Error GoTo tratar

    'Unload Me
   ' Dim rs As New ADODB.Recordset               'HOLDS ALL DATA RETURNED FROM QUERY

    'rs.CursorLocation = adUseClient
    'rs.Open SQL, MyDatabase, adOpenForwardOnly, adLockReadOnly, adCmdText
    'Set rs.ActiveConnection = Nothing
   '
    'If rs.EOF Then
    '    MsgBox "N�o h� registros!", vbInformation
    '    Unload Me
    '    Exit Sub
    'End If
          
    Dim crystal As New CRAXDRT.Application      'LOADS REPORT FROM FILE
    Dim report As CRAXDRT.report            'HOLDS REPORT
    Set report = crystal.OpenReport(PATH_REL & "RelVendasDiarias.rpt")  'OPEN OUR REPORT

    report.DiscardSavedData                      'CLEARS REPORT SO WE WORK FROM RECORDSET
    'report.Database.SetDataSource rs             'LINK REPORT TO RECORDSET
    report.RecordSelectionFormula = CRITERIO
    'report.GroupSelectionFormula = SelectionFormulaGrupo
    'formulas do relatorio
'    Dim i As Integer
'    Dim nomeFormula As String
'    For i = 1 To report.FormulaFields.Count
'        nomeFormula = UCase(report.FormulaFields.Item(i).Name)
'
'        If nomeFormula = "{@TITULO}" Then
'            report.FormulaFields.Item(i).Text = "'" + titulo + "'"
'        ElseIf nomeFormula = "{@FMLNOME1}" Then
'            report.FormulaFields.Item(i).Text = FMLNOME1
'        ElseIf nomeFormula = "{@FMLNOME2}" Then
'            report.FormulaFields.Item(i).Text = FMLNOME2
'        ElseIf nomeFormula = "{@FMLNOME3}" Then
'            report.FormulaFields.Item(i).Text = FMLNOME3
'        ElseIf nomeFormula = "{@PULAPAGINA1}" Then
'            report.FormulaFields.Item(i).Text = pulaPagina1
'        ElseIf nomeFormula = "{@PULAPAGINA2}" Then
'            report.FormulaFields.Item(i).Text = pulaPagina2
'        ElseIf nomeFormula = "{@PULAPAGINA3}" Then
'            report.FormulaFields.Item(i).Text = pulaPagina3
'        ElseIf nomeFormula = "{@TIPOANASINT}" Then
'            report.FormulaFields.Item(i).Text = TIPOANASINT
'        ElseIf nomeFormula = "{@SUBTITULO1}" Then
'            report.FormulaFields.Item(i).Text = SUBTITULO1
'        ElseIf nomeFormula = "{@SUBTITULO2}" Then
'            report.FormulaFields.Item(i).Text = SUBTITULO2
'        ElseIf nomeFormula = "{@SUBTITULO3}" Then
'            report.FormulaFields.Item(i).Text = SUBTITULO3
'        ElseIf nomeFormula = "{@TOTALGERAL}" Then
'            report.FormulaFields.Item(i).Text = TOTALGERAL
'        ElseIf nomeFormula = "{@PARAMETROS}" Then
'            report.FormulaFields.Item(i).Text = Parametros
'        ElseIf nomeFormula = "{@RODAPE}" Then
'            report.FormulaFields.Item(i).Text = RODAPE
'        ElseIf nomeFormula = "{@GENETICO1}" Then
'            report.FormulaFields.Item(i).Text = genetico1
'        ElseIf nomeFormula = "{@F1}" Then
'            report.FormulaFields.Item(i).Text = F1
'        ElseIf nomeFormula = "{@F2}" Then
'            report.FormulaFields.Item(i).Text = F2
'        ElseIf nomeFormula = "{@F3}" Then
'            report.FormulaFields.Item(i).Text = F3
'        ElseIf nomeFormula = "{@F4}" Then
'            report.FormulaFields.Item(i).Text = F4
'        ElseIf nomeFormula = "{@F5}" Then
'            report.FormulaFields.Item(i).Text = F5
'        'LIVRES
'        ElseIf nomeFormula = "{@FLIVRE1}" Then
'            report.FormulaFields.Item(i).Text = flivre1
'        ElseIf nomeFormula = "{@FLIVRE2}" Then
'            report.FormulaFields.Item(i).Text = flivre2
'        ElseIf nomeFormula = "{@FLIVRE3}" Then
'            report.FormulaFields.Item(i).Text = flivre3
'        ElseIf nomeFormula = "{@FLIVRE4}" Then
'            report.FormulaFields.Item(i).Text = flivre4
'        ElseIf nomeFormula = "{@FLIVRE5}" Then
'            report.FormulaFields.Item(i).Text = flivre5
'        End If
'
'    Next
'
'    Me.Show
    
    'CRViewer.DisplayBorder = False          'MAKES REPORT FILL ENTIRE FORM
    'CRViewer.DisplayTabs = False            'THIS REPORT DOES NOT DRILL DOWN, NOT NEEDED
    'CRViewer.EnableDrilldown = False        'REPORT DOES NOT SUPPORT DRILL-DOWN
    'CRViewer.EnableRefreshButton = False    'ADO RECORDSET WILL NOT CHANGE, NOT NEEDED
    'CRViewer.EnableExportButton = True      'EXPORTAR RELATORIO
    
    CRViewer.ReportSource = report              'LINK VIEWER TO REPORT
    CRViewer.ViewReport                   'SHOW REPORT

    Do While CRViewer.IsBusy              'ZOOM METHOD DOES NOT WORK WHILE
        DoEvents                          'REPORT IS LOADING, SO WE MUST PAUSE
    Loop                                  'WHILE REPORT LOADS.
        
    CRViewer.Top = 0                    'WHEN FORM IS RESIZED
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth

    'rs.Close
    'Set rs = Nothing
    Set crystal = Nothing
    Set report = Nothing
    
    'limparVariaveis
    Exit Sub
tratar:
    MsgBox Err.Description
End Sub

Public Sub SelecionaCampo(Formulario As Form, ByRef controle As Control)
On Error Resume Next
    
    If Len(controle) = 0 Then Exit Sub
    controle.SelStart = 0
    controle.SelLength = Len(controle)
End Sub

Public Function CalculaNossonumero(intCodigoContaCorrente As Integer, ByVal intSequenciaBoleto As String, Vencimento As Date) As String
    Dim Digito As String, resto As Integer
    Dim Calculo As Single
    Dim aux As String
    Dim rsNossoNumero As New ADODB.Recordset
    Dim strAgencia As String, strConta As String, booCobrancaComRegistro As Boolean
    Dim intBanco As Integer
    Dim lngSequenciaBoletoContaCorrente As Double
    Dim intConvenio As String
    Dim intCodigoCarteira As String
    Dim lngSequenciaBoletoFinal As Double
    Dim Condicao As Double
    If rsNossoNumero.State = 1 Then rsNossoNumero.Close
    rsNossoNumero.Open "SELECT CodigoCarteira, isnull(CodigoBanco,0) as banco, isnull(Agencia,'') as Agencia, isnull(Conta,'') as Conta, isnull(COMRegistro,1) as COMRegistro, isnull(SequenciaBoleta, 0) as SequenciaBoleta, Convenio, ISNULL(SequenciaBoletaFinal,0) AS SequenciaBoletaFinal FROM ContaCorrente WHERE Empresa_id = " & EMPRESA_ID_N & " AND CodigoContaCorrente = " & intCodigoContaCorrente, CONECTA_RETAGUARDA, , , adCmdText
    If Not rsNossoNumero.EOF Then
       intBanco = rsNossoNumero!Banco
       strAgencia = rsNossoNumero!Agencia
       strConta = rsNossoNumero!Conta
       booCobrancaComRegistro = rsNossoNumero!COMRegistro
       lngSequenciaBoletoContaCorrente = rsNossoNumero!SequenciaBoleta
       intConvenio = rsNossoNumero!convenio
       intCodigoCarteira = rsNossoNumero!CodigoCarteira
       lngSequenciaBoletoFinal = rsNossoNumero!SequenciaBoletaFinal
    Else
       MsgBox "Conta corrente inexistente...", 48, "Calcula nosso numero"
       rsNossoNumero.Close
       CalculaNossonumero = ""
       Exit Function
    End If
    rsNossoNumero.Close

    If lngSequenciaBoletoFinal > 0 Then
        If lngSequenciaBoletoContaCorrente >= lngSequenciaBoletoFinal Then
            MsgBox "Aten��o. Sequencia do boleto final foi superado, favor entrar em contato com o FINANCEIRO parar solicitar junto ao banco esta nova faixa num�rica.", vbCritical, "Calcula Nosso N�mero"
            CalculaNossonumero = ""
            Exit Function
        End If
        If lngSequenciaBoletoContaCorrente + 500 >= lngSequenciaBoletoFinal Then
            MsgBox "Aten��o. Sequencia do boleto final est� esgotando, favor entrar em contato com o FINANCEIRO parar solicitar junto ao banco esta nova faixa num�rica. Faltam apenas " & lngSequenciaBoletoFinal - lngSequenciaBoletoContaCorrente & " boletos para finalizar e ser bloqueado.", vbCritical, "Calcula Nosso N�mero"
        End If
    End If

    If intBanco = 1 Then 'banco do brasil
        CalculaNossonumero = lngSequenciaBoletoContaCorrente
        If Len(intConvenio) = 4 Or Len(intConvenio) = 6 Then
            If Len(intConvenio) = 4 Then
                aux = Trim$(Format(intConvenio, "0000")) & Trim$(Format(CalculaNossonumero, "0000000"))
            Else
                aux = Trim$(Format(intConvenio, "000000")) & Trim$(Format(CalculaNossonumero, "00000"))
            End If
            
            Calculo = 0
            Calculo = Val((Mid$(aux, 11, 1) * 9)) + Val((Mid$(aux, 10, 1) * 8)) + Val((Mid$(aux, 9, 1) * 7)) + Val((Mid$(aux, 8, 1) * 6)) + Val((Mid$(aux, 7, 1) * 5)) + Val((Mid$(aux, 6, 1) * 4)) + Val((Mid$(aux, 5, 1) * 3)) + Val((Mid$(aux, 4, 1) * 2)) + Val((Mid$(aux, 3, 1) * 9)) + Val((Mid$(aux, 2, 1) * 8)) + Val((Mid$(aux, 1, 1) * 7))
            resto = (Calculo Mod 11)
    
            If resto = 10 Then
                Digito = "X"
            ElseIf resto = 11 Then
                Digito = 0
            Else
                Digito = resto
            End If
            If Not IsNumeric(Digito) Then
                CalculaNossonumero = aux & Trim$(Digito)
            Else
                CalculaNossonumero = aux & Trim$(Str(Digito))
            End If
            
        ElseIf Len(intConvenio) = 7 Then
            aux = Trim$(Format(intConvenio, "0000000")) & Trim$(Format(CalculaNossonumero, "0000000000"))
            CalculaNossonumero = aux
            
            Calculo = 0
            Calculo = Val((Mid$(aux, 17, 1) * 9)) + Val((Mid$(aux, 16, 1) * 8)) + Val((Mid$(aux, 15, 1) * 7)) + Val((Mid$(aux, 14, 1) * 6)) + Val((Mid$(aux, 13, 1) * 5)) + Val((Mid$(aux, 12, 1) * 4)) + Val((Mid$(aux, 11, 1) * 3)) + Val((Mid$(aux, 10, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 11, 1) * 9)) + Val((Mid$(aux, 10, 1) * 8)) + Val((Mid$(aux, 9, 1) * 7)) + Val((Mid$(aux, 8, 1) * 6)) + Val((Mid$(aux, 7, 1) * 5)) + Val((Mid$(aux, 6, 1) * 4)) + Val((Mid$(aux, 5, 1) * 3)) + Val((Mid$(aux, 4, 1) * 2)) + Val((Mid$(aux, 3, 1) * 9)) + Val((Mid$(aux, 2, 1) * 8)) + Val((Mid$(aux, 1, 1) * 7))
            
            resto = (Calculo Mod 11)
    
            If resto = 10 Then
                Digito = "X"
            ElseIf resto = 11 Then
                Digito = 0
            Else
                Digito = resto
            End If
            If Not IsNumeric(Digito) Then
                CalculaNossonumero = aux & Trim$(Digito)
            Else
                CalculaNossonumero = aux & Trim$(Str(Digito))
            End If
        End If
        
    ElseIf intBanco = 237 Then 'banco bradesco
        CalculaNossonumero = lngSequenciaBoletoContaCorrente
        aux = Trim$(Format(CalculaNossonumero, "00000000000"))
        aux = Formata(CalculaNossonumero, 11)
        'sequencial
        Calculo = 0
        Calculo = Val((Mid$(aux, 11, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 10, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 9, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 8, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 7, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 6, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 5, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 4, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 3, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 2, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 1, 1) * 6))
    
        'carteira 09
        Calculo = Calculo + Val(intCodigoCarteira * 7)
        Calculo = Calculo + Val(0 * 2)
    
        resto = (Calculo Mod 11)
        Digito = 11 - resto
    
        If resto = 1 Then
            Digito = "P"
        ElseIf resto = 0 Then
            Digito = 0
        End If
    
        If Not IsNumeric(Digito) Then
            CalculaNossonumero = aux & Trim$(Digito)
        Else
            CalculaNossonumero = aux & Trim$((Digito))
        End If
        
    ElseIf intBanco = 479 Then 'banco boston
        CalculaNossonumero = lngSequenciaBoletoContaCorrente
        'CalculaNossonumero = 16054001
        aux = Trim$(Format(CalculaNossonumero, "00000000"))
        
        'sequencial
        Calculo = 0
        Calculo = Val((Mid$(aux, 8, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 7, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 6, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 5, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 4, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 3, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 2, 1) * 8))
        Calculo = Calculo + Val((Mid$(aux, 1, 1) * 9))
    
        Calculo = Calculo * 10
    
        resto = (Calculo Mod 11)
        Digito = resto
    
        If resto = 10 Then
            Digito = "0"
        End If
    
        CalculaNossonumero = aux & Trim$(Digito)
        
    ElseIf intBanco = 347 Then 'banco sudameris
        CalculaNossonumero = intSequenciaBoleto
        If booCobrancaComRegistro = True Then
            aux = Trim$(Format(intSequenciaBoleto, "0000000")) + Format(Left(strAgencia, 4), "0000") + Format(Left(strConta, 7), "0000000")
        Else
            aux = Trim$(Format(intSequenciaBoleto, "0000000000000")) + Format(Left(strAgencia, 4), "0000") + Format(Left(strConta, 7), "0000000")
        End If
        
        If booCobrancaComRegistro = True Then
            'sequencial
            Calculo = 0
            Calculo = Val((Mid$(aux, 18, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 17, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 16, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 15, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 14, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 13, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 12, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 11, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 10, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 9, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 8, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 7, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 6, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 5, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 4, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 3, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 2, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 1, 1) * 1))
        Else
            Calculo = 0
            Calculo = Val((Mid$(aux, 24, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 23, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 22, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 21, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 20, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 19, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 18, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 17, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 16, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 15, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 14, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 13, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 12, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 11, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 10, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 9, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 8, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 7, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 6, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 5, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 4, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 3, 1) * 1))
            Calculo = Calculo + Val((Mid$(aux, 2, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 1, 1) * 1))
        End If
        
        resto = (Calculo Mod 10)
        Digito = 10 - resto
        
        If resto = 10 Then
            Digito = "0"
        End If
        
        CalculaNossonumero = aux & Trim$(Digito)
        
    ElseIf intBanco = 422 Then 'banco SAFRA
        CalculaNossonumero = lngSequenciaBoletoContaCorrente
        aux = Trim$(Format(CalculaNossonumero, "00000000"))
        
        'sequencial
        Calculo = 0
        Calculo = Val((Mid$(aux, 8, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 7, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 6, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 5, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 4, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 3, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 2, 1) * 8))
        Calculo = Calculo + Val((Mid$(aux, 1, 1) * 9))
       
        resto = (Calculo Mod 11)
        If resto = 1 Then
            Digito = "0"
        ElseIf resto = 0 Then
            Digito = "1"
        Else
            Digito = 11 - resto
        End If
        
        CalculaNossonumero = aux & Trim$(Digito)
        
    ElseIf intBanco = 399 Then 'HSBC
        If booCobrancaComRegistro = True Then 'COBRANCA REGISTRADA
            CalculaNossonumero = lngSequenciaBoletoContaCorrente
            aux = Format(intConvenio, "00000") & Trim$(Format(CalculaNossonumero, "00000"))
            Calculo = 0
            Calculo = Val((Mid$(aux, 10, 1) * 2)) + Val((Mid$(aux, 9, 1) * 3)) + Val((Mid$(aux, 8, 1) * 4)) + Val((Mid$(aux, 7, 1) * 5)) + Val((Mid$(aux, 6, 1) * 6)) + Val((Mid$(aux, 5, 1) * 7)) + Val((Mid$(aux, 4, 1) * 2)) + Val((Mid$(aux, 3, 1) * 3)) + Val((Mid$(aux, 2, 1) * 4)) + Val((Mid$(aux, 1, 1) * 5))
            resto = (Calculo Mod 11)
            
            If resto = 0 Or resto = 1 Then
                Digito = "0"
            Else
                Digito = 11 - resto
            End If
            If Not IsNumeric(Digito) Then
                CalculaNossonumero = aux & Trim$(Digito)
            Else
                CalculaNossonumero = aux & Trim$(Str(Digito))
            End If
        Else 'COBRANCA NAO REGISTRADA
            Dim primeirodigito As String
            'CALCULA PRIMEIRO DIGITO
            CalculaNossonumero = lngSequenciaBoletoContaCorrente
            aux = Format(CalculaNossonumero, "00000000")
            Calculo = 0
            
            Calculo = Val((Mid$(aux, 8, 1) * 9)) + Val((Mid$(aux, 7, 1) * 8)) + Val((Mid$(aux, 6, 1) * 7)) + Val((Mid$(aux, 5, 1) * 6)) + Val((Mid$(aux, 4, 1) * 5)) + Val((Mid$(aux, 3, 1) * 4)) + Val((Mid$(aux, 2, 1) * 3)) + Val((Mid$(aux, 1, 1) * 2))
            resto = (Calculo Mod 11)
            
            If resto = 0 Or resto = 10 Then
                Digito = "0"
            Else
                Digito = resto
            End If
            aux = Format(CalculaNossonumero, "00000000") & Digito & "4"
            primeirodigito = aux
            
            Calculo = 0
            aux = Val(aux) + Val(Format(intConvenio, "0000000")) + Val(Format(Vencimento, "ddMMyy"))
            
            Calculo = Val((Mid$(aux, 10, 1) * 9)) + Val((Mid$(aux, 9, 1) * 8)) + Val((Mid$(aux, 8, 1) * 7)) + Val((Mid$(aux, 7, 1) * 6)) + Val((Mid$(aux, 6, 1) * 5)) + Val((Mid$(aux, 5, 1) * 4)) + Val((Mid$(aux, 4, 1) * 3)) + Val((Mid$(aux, 3, 1) * 2)) + Val((Mid$(aux, 2, 1) * 9)) + Val((Mid$(aux, 1, 1) * 8))
            resto = (Calculo Mod 11)
            
            If resto = 0 Or resto = 10 Then
                Digito = "0"
            End If
            If Not IsNumeric(Digito) Then
                CalculaNossonumero = primeirodigito & Trim$(Digito)
            Else
                CalculaNossonumero = primeirodigito & Trim$(Str(Digito))
            End If
        End If
        
    ElseIf intBanco = 341 Then 'banco ITAU
        CalculaNossonumero = intSequenciaBoleto
        aux = Format(Left(strAgencia, 4), "0000") + Format(Left(strConta, 5), "00000") + Format(intCodigoCarteira, "000") & Format(lngSequenciaBoletoContaCorrente, "00000000")
        Calculo = 0
        Calculo = IIf(Val((Mid$(aux, 20, 1) * 2)) < 10, Val((Mid$(aux, 20, 1) * 2)), Val(Left((Mid$(aux, 20, 1) * 2), 1)) + Val(Right((Mid$(aux, 20, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 19, 1) * 1)) < 10, Val((Mid$(aux, 19, 1) * 1)), Val(Left((Mid$(aux, 19, 1) * 1), 1)) + Val(Right((Mid$(aux, 19, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 18, 1) * 2)) < 10, Val((Mid$(aux, 18, 1) * 2)), Val(Left((Mid$(aux, 18, 1) * 2), 1)) + Val(Right((Mid$(aux, 18, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 17, 1) * 1)) < 10, Val((Mid$(aux, 17, 1) * 1)), Val(Left((Mid$(aux, 17, 1) * 1), 1)) + Val(Right((Mid$(aux, 17, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 16, 1) * 2)) < 10, Val((Mid$(aux, 16, 1) * 2)), Val(Left((Mid$(aux, 16, 1) * 2), 1)) + Val(Right((Mid$(aux, 16, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 15, 1) * 1)) < 10, Val((Mid$(aux, 15, 1) * 1)), Val(Left((Mid$(aux, 15, 1) * 1), 1)) + Val(Right((Mid$(aux, 15, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 14, 1) * 2)) < 10, Val((Mid$(aux, 14, 1) * 2)), Val(Left((Mid$(aux, 14, 1) * 2), 1)) + Val(Right((Mid$(aux, 14, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 13, 1) * 1)) < 10, Val((Mid$(aux, 13, 1) * 1)), Val(Left((Mid$(aux, 13, 1) * 1), 1)) + Val(Right((Mid$(aux, 13, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 12, 1) * 2)) < 10, Val((Mid$(aux, 12, 1) * 2)), Val(Left((Mid$(aux, 12, 1) * 2), 1)) + Val(Right((Mid$(aux, 12, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 11, 1) * 1)) < 10, Val((Mid$(aux, 11, 1) * 1)), Val(Left((Mid$(aux, 11, 1) * 1), 1)) + Val(Right((Mid$(aux, 11, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 10, 1) * 2)) < 10, Val((Mid$(aux, 10, 1) * 2)), Val(Left((Mid$(aux, 10, 1) * 2), 1)) + Val(Right((Mid$(aux, 10, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 9, 1) * 1)) < 10, Val((Mid$(aux, 9, 1) * 1)), Val(Left((Mid$(aux, 9, 1) * 1), 1)) + Val(Right((Mid$(aux, 9, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 8, 1) * 2)) < 10, Val((Mid$(aux, 8, 1) * 2)), Val(Left((Mid$(aux, 8, 1) * 2), 1)) + Val(Right((Mid$(aux, 8, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 7, 1) * 1)) < 10, Val((Mid$(aux, 7, 1) * 1)), Val(Left((Mid$(aux, 7, 1) * 1), 1)) + Val(Right((Mid$(aux, 7, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 6, 1) * 2)) < 10, Val((Mid$(aux, 6, 1) * 2)), Val(Left((Mid$(aux, 6, 1) * 2), 1)) + Val(Right((Mid$(aux, 6, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 5, 1) * 1)) < 10, Val((Mid$(aux, 5, 1) * 1)), Val(Left((Mid$(aux, 5, 1) * 1), 1)) + Val(Right((Mid$(aux, 5, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 4, 1) * 2)) < 10, Val((Mid$(aux, 4, 1) * 2)), Val(Left((Mid$(aux, 4, 1) * 2), 1)) + Val(Right((Mid$(aux, 4, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 3, 1) * 1)) < 10, Val((Mid$(aux, 3, 1) * 1)), Val(Left((Mid$(aux, 3, 1) * 1), 1)) + Val(Right((Mid$(aux, 3, 1) * 1), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 2, 1) * 2)) < 10, Val((Mid$(aux, 2, 1) * 2)), Val(Left((Mid$(aux, 2, 1) * 2), 1)) + Val(Right((Mid$(aux, 2, 1) * 2), 1)))
        Calculo = Calculo + IIf(Val((Mid$(aux, 1, 1) * 1)) < 10, Val((Mid$(aux, 1, 1) * 1)), Val(Left((Mid$(aux, 1, 1) * 1), 1)) + Val(Right((Mid$(aux, 1, 1) * 1), 1)))
        
        
    '            CALCULO = CALCULO + Val((Mid$(Aux, 19, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 18, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 17, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 16, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 15, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 14, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 13, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 12, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 11, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 10, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 9, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 8, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 7, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 6, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 5, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 4, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 3, 1) * 1))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 2, 1) * 2))
    '            CALCULO = CALCULO + Val((Mid$(Aux, 1, 1) * 1))
        
        resto = (Calculo Mod 10)
        
        If resto = 0 Then
            Digito = "0"
        Else
            Digito = 10 - resto
        End If
        
        CalculaNossonumero = Format(lngSequenciaBoletoContaCorrente, "00000000") & Trim$(Str(Digito))
    
    ElseIf intBanco = 356 Then 'banco real
        CalculaNossonumero = lngSequenciaBoletoContaCorrente
    ElseIf intBanco = 320 Then 'bicbanco
    
        CalculaNossonumero = lngSequenciaBoletoContaCorrente
        aux = Format(strAgencia, "000") & "" & Trim$(Format(CalculaNossonumero, "000000"))
        aux = Formata(aux, 9)
        'sequencial
        Calculo = 0
        Calculo = Val((Mid$(aux, 9, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 8, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 7, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 6, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 5, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 4, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 3, 1) * 8))
        Calculo = Calculo + Val((Mid$(aux, 2, 1) * 9))
        Calculo = Calculo + Val((Mid$(aux, 1, 1) * 2))
    
        resto = (Calculo Mod 11)
        Digito = 11 - resto
    
        If resto = 1 Then
            Digito = "0"
        ElseIf resto = 0 Then
            Digito = 1
        End If
    
        If Not IsNumeric(Digito) Then
            CalculaNossonumero = aux & Trim$(Digito)
        Else
            CalculaNossonumero = aux & Trim$((Digito))
        End If
        
    ElseIf intBanco = 70 Then 'BRB
        
        CalculaNossonumero = lngSequenciaBoletoContaCorrente
        'Calculo com 23 posicoes para Digito (D1), zeros(3),Agencia(3),Conta(7),Categoria(1),Sequencial(6),Banco(3)
        aux = "000" & Format(Left(strAgencia, 3), "000") & Format(Left(strConta, 7), "0000000") & Format(intCodigoCarteira, "0") & Format(lngSequenciaBoletoContaCorrente, "000000") & "070"
        'sequencial
        Calculo = 0
        
        'Se a Multiplicacao do Produto for > 9  o produto e diminuido por 9 conforme manual do BRB
        If Val((Mid$(aux, 23, 1) * 2)) > 9 Then
           Calculo = Val((Mid$(aux, 13, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 23, 1) * 2)) < 10 Then
           Calculo = Val((Mid$(aux, 23, 1) * 2))
        End If
        
        If Val((Mid$(aux, 22, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 22, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 22, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 22, 1) * 1))
        End If
        
        If Val((Mid$(aux, 21, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 21, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 21, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 21, 1) * 2))
        End If
        
        If Val((Mid$(aux, 20, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 20, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 20, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 20, 1) * 1))
        End If
        
        If Val((Mid$(aux, 19, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 19, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 19, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 19, 1) * 2))
        End If
        
        If Val((Mid$(aux, 18, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 18, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 18, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 18, 1) * 1))
        End If
        
        If Val((Mid$(aux, 17, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 17, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 17, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 17, 1) * 2))
        End If
        
        If Val((Mid$(aux, 16, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 16, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 16, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 16, 1) * 1))
        End If
        
        If Val((Mid$(aux, 15, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 15, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 15, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 15, 1) * 2))
        End If
        
        If Val((Mid$(aux, 14, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 14, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 14, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 14, 1) * 1))
        End If
        
        If Val((Mid$(aux, 13, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 13, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 13, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 13, 1) * 2))
        End If
           
        If Val((Mid$(aux, 12, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 12, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 12, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 12, 1) * 1))
        End If
        
        If Val((Mid$(aux, 11, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 11, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 11, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 11, 1) * 2))
        End If
        
        If Val((Mid$(aux, 10, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 10, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 10, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 10, 1) * 1))
        End If
        
        If Val((Mid$(aux, 9, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 9, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 9, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 9, 1) * 2))
        End If
        
        If Val((Mid$(aux, 8, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 8, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 8, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 8, 1) * 1))
        End If
        
        If Val((Mid$(aux, 7, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 7, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 7, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 7, 1) * 2))
        End If
        
        If Val((Mid$(aux, 6, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 6, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 6, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 6, 1) * 1))
        End If
        
        If Val((Mid$(aux, 5, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 5, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 5, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 5, 1) * 2))
        End If
        
        If Val((Mid$(aux, 4, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 4, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 4, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 4, 1) * 1))
        End If
        
        If Val((Mid$(aux, 3, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 3, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 3, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 3, 1) * 2))
        End If
        
        If Val((Mid$(aux, 2, 1) * 1)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 2, 1) * 1)) - 9
        ElseIf Val((Mid$(aux, 2, 1) * 1)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 2, 1) * 1))
        End If
        
        If Val((Mid$(aux, 1, 1) * 2)) > 9 Then
           Calculo = Calculo + Val((Mid$(aux, 1, 1) * 2)) - 9
        ElseIf Val((Mid$(aux, 1, 1) * 2)) < 10 Then
           Calculo = Calculo + Val((Mid$(aux, 1, 1) * 2))
        End If
        
        'Calculo do digito D1
        resto = (Calculo Mod 10)
        If resto > 0 Then
           Digito = 10 - resto
        ElseIf resto = 0 Then
           Digito = 0
        End If
        aux = aux & Digito
RecalculaDigitoD2:
        
        
        'Calculo com 24 posicoes para Digito (D2), zeros(3),Agencia(3),Conta(7),Categoria(1),Sequencial(6),Banco(3), Digito1(D1)
        Calculo = 0
        Calculo = Val((Mid$(aux, 24, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 23, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 22, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 21, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 20, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 19, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 18, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 17, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 16, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 15, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 14, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 13, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 12, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 11, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 10, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 9, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 8, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 7, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 6, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 5, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 4, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 3, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 2, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 1, 1) * 7))
        
        'Calculo do digito D2
        resto = (Calculo Mod 11)
        If resto > 1 Then
           Digito = 11 - resto
        ElseIf resto = 0 Then
           Digito = 0
        ElseIf resto = 1 Then
           'Neste Caso aqui se o resto for = 1 o Digito 2 (D2) Sera Recalculado com o novo digito 1 (D1)
           'Conforme Manual BRB
           Digito = 1 + Mid(aux, 24, 1)
           If Digito = 10 Then
              Digito = 0
              aux = "000" & Format(Left(strAgencia, 3), "000") & Format(Left(strConta, 7), "0000000") & Format(intCodigoCarteira, "0") & Format(lngSequenciaBoletoContaCorrente, "000000") & "070"
              aux = aux & Digito
              GoTo RecalculaDigitoD2
           ElseIf Digito <> 10 Then
              aux = "000" & Format(Left(strAgencia, 3), "000") & Format(Left(strConta, 7), "0000000") & Format(intCodigoCarteira, "0") & Format(lngSequenciaBoletoContaCorrente, "000000") & "070"
              aux = aux & Digito
              GoTo RecalculaDigitoD2
           End If
           
        End If
        aux = aux & Digito
        aux = Mid(aux, 14, 12) 'Pegando o Nosso Numero de tamanho 12 gerado com os dois digitos conforme Manual
        
        If Len(aux) = 12 Then
            CalculaNossonumero = aux
        End If
        
    ElseIf intBanco = 33 Or intBanco = 353 Then 'SANTANDER ficou no lugar do banespa
        CONECTA_RETAGUARDA.Execute "UPDATE Banco SET Descricao = 'SANTANDER' WHERE CodigoBanco = 33 and Descricao = 'BANESPA'"
        
        '0000001 a 9999999
        CalculaNossonumero = Format(lngSequenciaBoletoContaCorrente, "000000000000")
        aux = CalculaNossonumero
        
        'sequencial
        Calculo = 0
        Calculo = Val((Mid$(aux, 12, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 11, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 10, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 9, 1) * 5))
        Calculo = Calculo + Val((Mid$(aux, 8, 1) * 6))
        Calculo = Calculo + Val((Mid$(aux, 7, 1) * 7))
        Calculo = Calculo + Val((Mid$(aux, 6, 1) * 8))
        Calculo = Calculo + Val((Mid$(aux, 5, 1) * 9))
        Calculo = Calculo + Val((Mid$(aux, 4, 1) * 2))
        Calculo = Calculo + Val((Mid$(aux, 3, 1) * 3))
        Calculo = Calculo + Val((Mid$(aux, 2, 1) * 4))
        Calculo = Calculo + Val((Mid$(aux, 1, 1) * 5))
    
        
        resto = (Calculo Mod 11)
        
    
        If resto = 10 Then
            Digito = 1
        ElseIf resto = 0 Or resto = 1 Then
            Digito = 0
        Else
            Digito = 11 - resto
        End If
        
    
        If Not IsNumeric(Digito) Then
            CalculaNossonumero = aux & Trim$(Digito)
        Else
            CalculaNossonumero = aux & Trim$((Digito))
        End If
      
    ElseIf intBanco = 104 Then 'CAIXA ECON�MICA FEDERAL CEF
    
    '            8.2.1 Para a carteira 11 - Cobran�a Simples (Vide Nota 3): N�mero gerado e atribu�do pelo sistema de cobran�a da
    '            CAIXA para controle interno, e ser� composto da seguinte forma:
    '            NNNNNNNNNND , onde
    '            NNNNNNNNNN = N�mero Sequencial
    '            D = D�gito Verificador (calculado pelo Mod. 11)
    '            Obs: para clientes que possuem sistema pr�prio, preencher o campo com zeros.
    '            8.2.2 Para a carteira 12 - Cobran�a R�pida: N�mero informado pelo cliente, composto da seguinte forma:
    '            9NNNNNNNNND, onde 9 = Fixo
    '            NNNNNNNNN = N�mero Sequencial
    '            D = D�gito Verificador (calculado pelo Mod. 11)
    
        If intCodigoCarteira = 11 Then 'COBRAN�A SIMPLES
            CalculaNossonumero = Format(lngSequenciaBoletoContaCorrente, "0000000000")
            aux = CalculaNossonumero
        
            'sequencial
            Calculo = 0
            Calculo = Val((Mid$(aux, 10, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 9, 1) * 3))
            Calculo = Calculo + Val((Mid$(aux, 8, 1) * 4))
            Calculo = Calculo + Val((Mid$(aux, 7, 1) * 5))
            Calculo = Calculo + Val((Mid$(aux, 6, 1) * 6))
            Calculo = Calculo + Val((Mid$(aux, 5, 1) * 7))
            Calculo = Calculo + Val((Mid$(aux, 4, 1) * 8))
            Calculo = Calculo + Val((Mid$(aux, 3, 1) * 9))
            Calculo = Calculo + Val((Mid$(aux, 2, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 1, 1) * 3))
            
            resto = (Calculo Mod 11)
            Digito = 11 - resto
    
            If Digito > 9 Then
               Digito = 0
            End If
            
        ElseIf intCodigoCarteira = 12 Then 'COBRAN�A R�PIDA
            CalculaNossonumero = "9" & Format(lngSequenciaBoletoContaCorrente, "000000000")
            
        ElseIf intCodigoCarteira = 14 Then 'COBRAN�A REGISTRADA COM CEDENTE
            'CalculaNossonumero = "82" & format(lngSequenciaBoletoContaCorrente, "0000000000000")
            CalculaNossonumero = "14" & Format(lngSequenciaBoletoContaCorrente, "000000000000000")
            aux = CalculaNossonumero
            'sequencial
            Calculo = 0
            Calculo = Val((Mid$(aux, 17, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 16, 1) * 3))
            Calculo = Calculo + Val((Mid$(aux, 15, 1) * 4))
            Calculo = Calculo + Val((Mid$(aux, 14, 1) * 5))
            Calculo = Calculo + Val((Mid$(aux, 13, 1) * 6))
            Calculo = Calculo + Val((Mid$(aux, 12, 1) * 7))
            Calculo = Calculo + Val((Mid$(aux, 11, 1) * 8))
            Calculo = Calculo + Val((Mid$(aux, 10, 1) * 9))
            Calculo = Calculo + Val((Mid$(aux, 9, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 8, 1) * 3))
            Calculo = Calculo + Val((Mid$(aux, 7, 1) * 4))
            Calculo = Calculo + Val((Mid$(aux, 6, 1) * 5))
            Calculo = Calculo + Val((Mid$(aux, 5, 1) * 6))
            Calculo = Calculo + Val((Mid$(aux, 4, 1) * 7))
            Calculo = Calculo + Val((Mid$(aux, 3, 1) * 8))
            Calculo = Calculo + Val((Mid$(aux, 2, 1) * 9))
            Calculo = Calculo + Val((Mid$(aux, 1, 1) * 2))
            If Calculo < 11 Then
               resto = Calculo
               Digito = 11 - resto
            Else
               resto = (Calculo Mod 11)
               Digito = 11 - resto
               If Digito > 9 Then
                  Digito = 0
               End If
            End If
        ElseIf intCodigoCarteira = 24 Then 'COBRAN�A SEM REGISTRO COM CEDENTE
            'CalculaNossonumero = "82" & format(lngSequenciaBoletoContaCorrente, "0000000000000")
            CalculaNossonumero = "24" & Format(lngSequenciaBoletoContaCorrente, "000000000000000")
            aux = CalculaNossonumero
            'sequencial
            Calculo = 0
            Calculo = Val((Mid$(aux, 17, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 16, 1) * 3))
            Calculo = Calculo + Val((Mid$(aux, 15, 1) * 4))
            Calculo = Calculo + Val((Mid$(aux, 14, 1) * 5))
            Calculo = Calculo + Val((Mid$(aux, 13, 1) * 6))
            Calculo = Calculo + Val((Mid$(aux, 12, 1) * 7))
            Calculo = Calculo + Val((Mid$(aux, 11, 1) * 8))
            Calculo = Calculo + Val((Mid$(aux, 10, 1) * 9))
            Calculo = Calculo + Val((Mid$(aux, 9, 1) * 2))
            Calculo = Calculo + Val((Mid$(aux, 8, 1) * 3))
            Calculo = Calculo + Val((Mid$(aux, 7, 1) * 4))
            Calculo = Calculo + Val((Mid$(aux, 6, 1) * 5))
            Calculo = Calculo + Val((Mid$(aux, 5, 1) * 6))
            Calculo = Calculo + Val((Mid$(aux, 4, 1) * 7))
            Calculo = Calculo + Val((Mid$(aux, 3, 1) * 8))
            Calculo = Calculo + Val((Mid$(aux, 2, 1) * 9))
            Calculo = Calculo + Val((Mid$(aux, 1, 1) * 2))
            If Calculo < 11 Then
               resto = Calculo
               Digito = 11 - resto
            Else
               resto = (Calculo Mod 11)
               Digito = 11 - resto
               If Digito > 9 Then
                  Digito = 0
               End If
            End If
        ElseIf intCodigoCarteira = 41 Then 'DESCONTADO
        
        End If


        If Not IsNumeric(Digito) Then
            CalculaNossonumero = aux & Trim$(Digito)
        Else
            CalculaNossonumero = aux & Trim$((Digito))
        End If
    End If

End Function

Public Function Formata(ByVal Valor As String, ByVal tamanho As Integer) As String
    Dim strAux As String
    strAux = Mid("              ", 1, tamanho - Len(Valor)) & Valor
    Dim strConcatena As String
    Dim intTamanhoPercorrido As Integer
    strConcatena = ""
    For intTamanhoPercorrido = 1 To tamanho
        If Mid(strAux, intTamanhoPercorrido, 1) = " " Then
            strConcatena = strConcatena & "0"
        Else
            strConcatena = strConcatena & Mid(strAux, intTamanhoPercorrido, 1)
        End If
    Next intTamanhoPercorrido
    Formata = strConcatena
End Function

Public Sub Tempo(ByVal iSegundos As Integer)
'On Error GoTo ERRO_TRATA

    Dim vInicio As Variant
    
    vInicio = Time
    While DateDiff("s", vInicio, Time) < iSegundos
    Wend
    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descri��o: " & Err.Description & " Origem:" & Err.Source, 48, "Tempo"
    End If
End Sub

Public Function BuscaCodigo(SQL As String, Banco As Variant, Campo As Variant) As String
'On Error GoTo ERRO_TRATA

    Dim rs_Maior As New ADODB.Recordset
    
    BuscaCodigo = ""
    rs_Maior.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
    If Not rs_Maior.EOF Then
        BuscaCodigo = IIf(Not IsNull(rs_Maior(Campo)), rs_Maior(Campo), 0)
    Else
        BuscaCodigo = 0
    End If
    rs_Maior.Close
    
    Exit Function
ERRO_TRATA:
'    ControleErros Err.Number, Err.Description, Err.Source, "codigoCombo"
End Function

Public Function fnAlinhaCampos(Valor As String, tamanho As Integer, Ajuste As String, Caracter As String)
   Dim Concatena  As String
   Dim GrupoCaracter As String
   GrupoCaracter = ""
   
   For i = 1 To tamanho
       GrupoCaracter = GrupoCaracter & Caracter
   Next i


   Concatena = ""
   If Ajuste = "E" Then
       Concatena = Format(Valor * 100, GrupoCaracter)
       'Concatena = "" 'Replace(Caracter, Tamanho - Len(Valor)) + Trim(Valor)
   Else
       Concatena = Format(Valor * 100, GrupoCaracter)
       'Concatena = "" ''RTrim(Valor) + Replace(Caracter, Tamanho - Len(Valor))"
   End If
   If Len(Valor) > tamanho Then Concatena = Left(Trim(Valor), tamanho)
   fnAlinhaCampos = Concatena
End Function

Public Function SomenteNumeros(Enter As Integer) As Integer
    If Enter = 8 Or Enter = 13 Or Enter = 45 Then 'Se for DELETE ou Backspace n�o fazer nada
        SomenteNumeros = Enter
        Exit Function
    End If
    
    If Enter = 46 Or Enter = 44 Then  'Se for ponto ou v�rcula assume v�rgula
        SomenteNumeros = 44
        Exit Function
    End If
    
    If Not IsNumeric(Chr(Enter)) Then  'Se n�o for n�mero retorne zero
        SomenteNumeros = 0
        Exit Function
    End If
    
    If Enter = 8 Then                 'Se for Delete ou Backspace n�o fazer nada
        SomenteNumeros = Enter
        Exit Function
    End If
    
    If Enter < 48 Or Enter > 57 Then  'Se n�o for n�mero retorne zero
        SomenteNumeros = 0
        Exit Function
    End If
    
    SomenteNumeros = Enter
End Function

Public Function FimDoMes(strData As String, blnSaltaMesAtual As Boolean) As String
   Dim strAno        As String
   Dim strMes        As String
   Dim strDia        As String
   Dim strProximoDia As String

   strData = Format(strData, "yyyymmdd")

   strAno = Mid$(strData, 1, 4)
   strMes = Mid$(strData, 5, 2)

   ' Pega a data e o mes atual

   Select Case strMes
      Case "04", "06", "09", "11"
      strDia = "30"

      Case "02"
         If Bissexto(Val(strAno)) Then
            strDia = "29"
            Else: strDia = "28"
         End If

      Case Else
         strDia = "31"
   End Select

   FimDoMes = strAno & strMes & strDia

   If (FimDoMes = strData) And SaltaMesAtual Then
      strProximoDia = ProximoDia(strData)
      FimDoMes = FimDoMes(strProximoDia, False)
   End If
End Function

Public Function ProximoDia(strDatea As String) As String
   Dim dteData As Date
   'Converte a data para o formato "yyyymmdd"
   dteData = Format(strData, "@@@@-@@-@@")
   ProximoDia = Format$(DateAdd("d", 1, dteData), "yyyymmdd")
End Function

Public Function BUSCA_TRIBUTA��O_PRODUTO(strSituacao_Tributaria_Produto As String) As String
'On Error GoTo ERRO_TRATA

'tentando fazer a PORRA FUNCIONAR
'101 � Tributada pelo Simples Nacional com permiss�o de cr�dito
'102 � Tributada pelo Simples Nacional sem permiss�o de cr�dito
'103 � Isen��o do ICMS no Simples Nacional para faixa de receita bruta
'201 � Tributada pelo Simples Nacional com permiss�o de cr�dito e com cobran�a do ICMS por substitui��o tribut�ria
'202 � Tributada pelo Simples Nacional sem permiss�o de cr�dito e com cobran�a do ICMS por substitui��o tribut�ria
'203 � Isen��o do ICMS no Simples Nacional para a faixa de receita bruta e com cobran�a de ICMS por substitui��o tribut�ria
'300 � Imune
'400 � N�o tributada pelo Simples Nacional
'500 � ICMS cobrado anteriormente por substitui��o tribut�ria
'900 � Outros.

   intTributacao = 400  'N�o tributada pelo Simples Nacional

   'Tributada integralmente = 00
   If strSituacao_Tributaria_Produto = "00" Then
      '101 � Tributada pelo Simples Nacional com permiss�o de cr�dito
      '102 � Tributada pelo Simples Nacional sem permiss�o de cr�dito
      intTributacao = 101
   End If

   'Tributada e com cobran�a do ICMS por substitui��o tribut�ria = 10
   If strSituacao_Tributaria_Produto = "10" Then
      '201 � Tributada pelo Simples Nacional com permiss�o de cr�dito e com cobran�a do ICMS por substitui��o tribut�ria
      '202 � Tributada pelo Simples Nacional sem permiss�o de cr�dito e com cobran�a do ICMS por substitui��o tribut�ria
      intTributacao = 201
   End If

   'Com redu��o de base de c�lculo = 20
   If strSituacao_Tributaria_Produto = "20" Then
      'intTributacao =
   End If

   'Isenta ou n�o tributada e com cobran�a do ICMS por substitui��o tribut�ria = 30
   If strSituacao_Tributaria_Produto = "30" Then
      '103 � Isen��o do ICMS no Simples Nacional para faixa de receita bruta
      intTributacao = 103
   End If

   'Isenta = 40
   If strSituacao_Tributaria_Produto = "40" Then
      '400 � N�o tributada pelo Simples Nacional
      intTributacao = 400
   End If

   'N�o tributada = 41
   If strSituacao_Tributaria_Produto = "41" Then
      '400 � N�o tributada pelo Simples Nacional
      intTributacao = 400
   End If

   'Suspens�o = 50
   If strSituacao_Tributaria_Produto = "50" Then
      '300 � Imune
      intTributacao = 300
   End If

   'Diferimento = 51
   If strSituacao_Tributaria_Produto = "51" Then _
      intTributacao = 400

   'ICMS cobrado anteriormente por substitui��o tribut�ria = 60
   If strSituacao_Tributaria_Produto = "60" Then
      '500 � ICMS cobrado anteriormente por substitui��o tribut�ria
      intTributacao = 500
   End If

   'Com redu��o de base de c�lculo e cobran�a de ICMS por substitui��o tribut�ria = 70
   '203 � Isen��o do ICMS no Simples Nacional para a faixa de receita bruta e com cobran�a de ICMS por substitui��o tribut�ria

   'Outras
   '900 � Outros.

'esta fixo aqui porque ainda falta entendimento se a regra acima procede
If intTributacao <> 500 Then _
   intTributacao = 400  'N�o tributada pelo Simples Nacional

   BUSCA_TRIBUTA��O_PRODUTO = intTributacao

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "BUSCA_TRIBUTA��O_PRODUTO", "BUSCA_TRIBUTA��O_PRODUTO"
End Function

Public Function Checagem_Defini��o_Campo_Tabela(strTabela As String, _
                                                strCampo As String, _
                                                strTipoCampo As String, _
                                                strOperacao As Integer) As Boolean
'On Error GoTo ERRO_TRATA

   If Trim(strTipoCampo) <> "" Then
      If ExisteTabela("RETAGUARDA", strTabela, "") = False Then
         If ExisteCampo("RETAGUARDA", strCampo, strTabela) = False Then
            'altera��o campo
            If strOperacao = 1 Then _
               CONECTA_RETAGUARDA.Execute "ALTER " & strTabela & " ALTER COLUMN " & strCampo & " " & strTipoCampo
            'criar campo
            If strOperacao = 2 Then _
               CONECTA_RETAGUARDA.Execute "ALTER " & strTabela & " ADD COLUMN " & strCampo & " " & strTipoCampo
         End If
      End If
   End If

Exit Function
ERRO_TRATA:
   Checagem_Defini��o_Campo_Tabela = False
End Function

Public Sub Altera��o_Defini��o_Campo_Tabela(strCampo As String, strTipoCampoAlter As String, strTabela As String)
On Error Resume Next

   Dim strVACA As String

   If Trim(strCampo) = "" Then _
      Exit Sub

   If Trim(strTipoCampoAlter) = "" Then _
      Exit Sub

   CONT_N = 0

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   strVACA = ""

   strVACA = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS"
   strVACA = strVACA & " WHERE COLUMN_NAME = '" & Trim(UCase(strCampo)) & "'"
   If Trim(strTabela) <> "" Then _
      strVACA = strVACA & " and table_NAME = '" & Trim(UCase(strTabela)) & "'"
   TabConsulta.Open strVACA, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabConsulta.EOF
      If Trim(UCase(TabConsulta.Fields("data_type").Value)) <> Trim(UCase(strTipoCampoAlter)) Then
         SQL = "ALTER TABLE " & Trim(TabConsulta.Fields("table_name").Value) & " ALTER COLUMN " & strCampo & " " & Trim(strTipoCampoAlter)

         CONECTA_RETAGUARDA.Execute SQL
         CONT_N = CONT_N + 1
      End If

      TabConsulta.MoveNext
      Err.Clear
   Wend

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "Altera��o_Defini��o_Campo_Tabela"
End Sub

Sub ESCOLHE_IMPRESSORA(Banco_Dados As String)
   If Trim(UCase(Banco_Dados)) = UCase("GLOBAL") Then
      Crystaldsn = Servidor_Global
      Crystaldsq = Banco_Dados
      Crystaluid = USUARIO_ADM_SQLSERVER
      Crystalpwd = SENHA_ADM_SQLSERVER
      Else
         Crystaldsn = SERVIDOR_SHFSYS
         Crystaldsq = Banco_Dados
         Crystaluid = USUARIO_ADM_SQLSERVER
         Crystalpwd = SENHA_ADM_SQLSERVER
   End If

   frmINICIO.Dialogo.CancelError = True
   frmINICIO.Dialogo.ShowPrinter
End Sub

Sub TABELAS_RETAGUARDA()
'============TABELA PESSOA
   If ExisteTabela("RETAGUARDA", "PESSOA", "") = False Then
      SQL = "CREATE TABLE PESSOA("

         SQL = SQL & " PESSOA_ID bigint NOT NULL,"
         SQL = SQL & " CNPJCPF NVARCHAR(14) NULL,"
         SQL = SQL & " DESCRICAO NVARCHAR(max) NOT NULL,"
         SQL = SQL & " RAZAO varchar(max) NULL,"
         SQL = SQL & " DATA_CAD datetime NOT NULL,"
         SQL = SQL & " SITUACAO NVARCHAR(1) NOT NULL,"

      SQL = SQL & " CONSTRAINT PK_PESSOA PRIMARY KEY CLUSTERED"
      SQL = SQL & " (PESSOA_ID Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

'============TABELA EMPRESA
   If ExisteTabela("RETAGUARDA", "EMPRESA", "") = False Then
      SQL = "CREATE TABLE EMPRESA("

         SQL = SQL & " EMPRESA_ID bigint NOT NULL,"
         SQL = SQL & " PESSOA_ID bigint NOT NULL,"
         SQL = SQL & " SITUACAO NVARCHAR(1) NOT NULL,"

      SQL = SQL & " CONSTRAINT PK_PESSOA PRIMARY KEY CLUSTERED"
      SQL = SQL & " (PESSOA_ID Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

'============TABELA USUARIO
   If ExisteTabela("RETAGUARDA", "USUARIO", "") = False Then
      SQL = SQL & " CREATE TABLE USUARIO("
      SQL = SQL & " USUARIO_ID bigint NOT NULL,"
      SQL = SQL & " PESSOA_ID bigint not NULL,"
      SQL = SQL & " LOGON nchar(30) not NULL,"
      SQL = SQL & " SENHA nchar(30) not NULL,"
      SQL = SQL & " SITUACAO NVARCHAR(50) not NULL,"
      SQL = SQL & " CONSTRAINT PK_USUARIO PRIMARY KEY CLUSTERED"
      SQL = SQL & " (USUARIO_ID Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
   End If
'==================================================

   If ExisteTabela("RETAGUARDA", "PRODUTO", "") = True Then _
      If ExisteCampo("RETAGUARDA", "STATUS", "PRODUTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN STATUS"

   If ExisteTabela("RETAGUARDA", "PEDIDOITEM", "") = True Then
      If ExisteCampo("RETAGUARDA", "SEQ", "PEDIDOITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM DROP COLUMN SEQ"

      'If ExisteCampo("RETAGUARDA","EMPRESA_ID", "PEDIDOITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM DROP COLUMN EMPRESA_ID"
   End If

   If ExisteTabela("RETAGUARDA", "PRODUTO", "") = True Then
      If ExisteCampo("RETAGUARDA", "QTDE", "PRODUTO") = True Then
         Altera��o_Defini��o_Campo_Tabela "QTDE", "NUMERIC(18,3)", "PRODUTO"
         Altera��o_Defini��o_Campo_Tabela "qtde_retido", "NUMERIC(18,3)", "PRODUTO"
      End If
   End If

   If ExisteTabela("RETAGUARDA", "PEDIDOITEM", "") = True Then
      If ExisteCampo("RETAGUARDA", "PRECO_CUSTO", "PEDIDOITEM") = False Then
         SQL = "ALTER TABLE PEDIDOITEM ADD PRECO_CUSTO float"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If ExisteTabela("RETAGUARDA", "CAIXADIA", "") = True Then
      If ExisteCampo("RETAGUARDA", "VALOR_DOLAR", "CAIXADIA") = False Then
         SQL = "ALTER TABLE CAIXADIA ADD VALOR_DOLAR float"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If ExisteTabela("RETAGUARDA", "PEDIDO", "") = True Then
      If ExisteCampo("RETAGUARDA", "NUMR_CUPOM", "PEDIDO") = False Then
         SQL = "ALTER TABLE PEDIDO ADD NUMR_CUPOM NVARCHAR(20)"
         CONECTA_RETAGUARDA.Execute SQL
         Else: Altera��o_Defini��o_Campo_Tabela "NUMR_CUPOM", "NVARCHAR(20)", "PEDIDO"
      End If
   End If
'==============================================
   If ExisteTabela("RETAGUARDA", "FONE", "") = True Then _
      If ExisteCampo("RETAGUARDA", "EMPRESA_ID", "FONE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FONE DROP COLUMN EMPRESA_ID"

   If ExisteTabela("RETAGUARDA", "BANCO", "") = True Then _
      If ExisteCampo("RETAGUARDA", "EMPRESA_ID", "BANCO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE BANCO DROP COLUMN EMPRESA_ID"

   If ExisteTabela("RETAGUARDA", "AGENCIA", "") = False Then
      SQL = "CREATE TABLE [dbo].[AGENCIA]("
      SQL = SQL & " [NUMR_AGENCIA] [varchar](10) NOT NULL,"
      SQL = SQL & " [BANCO] [int] NOT NULL,"
      SQL = SQL & " [NOME_AGENCIA] [varchar](100) NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If ExisteCampo("RETAGUARDA", "EMPRESA_ID", "AGENCIA") = True Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE AGENCIA DROP COLUMN EMPRESA_ID"
   End If

   If ExisteTabela("RETAGUARDA", "CONTA", "") = True Then
      If ExisteCampo("RETAGUARDA", "PESSOA_ID", "CONTA") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA ADD PESSOA_ID BIGINT"

         SQL = "update CONTA set pessoa_id = codg_cliente"
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA DROP COLUMN CODG_CLIENTE"
      End If
      If ExisteCampo("RETAGUARDA", "EMPRESA_ID", "CONTA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA DROP COLUMN EMPRESA_ID"
   End If

   If ExisteTabela("RETAGUARDA", "CHEQUE", "") = True Then
      If ExisteCampo("RETAGUARDA", "EMPRESA_ID", "CHEQUE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE DROP COLUMN EMPRESA_ID"
   
      If ExisteCampo("RETAGUARDA", "NUMR_DOC", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD NUMR_DOC NVARCHAR(30)"
   End If
'==============================================
   If ExisteTabela("RETAGUARDA", "PEDIDOITEM", "") = True Then
      If ExisteCampo("RETAGUARDA", "STATUS", "PEDIDOITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD STATUS CHAR(1)"
      End If
   End If

'==================================================
   If ExisteTabela("RETAGUARDA", "CSON", "") = False Then
      SQL = " CREATE TABLE CSON("
      SQL = SQL & " CODIGO NVARCHAR(3) NOT NULL,"
      SQL = SQL & " DESCRICAO NVARCHAR(max) NOT NULL,"
      SQL = SQL & " OBS NVARCHAR(max) NULL,"
      SQL = SQL & " CONSTRAINT PK_TRIBUTACAO_CSON PRIMARY KEY CLUSTERED"
      SQL = SQL & " (Codigo Asc )"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 101
      SQL = SQL & ",'Tributada pelo Simples Nacional com permiss�o de cr�dito'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es que permitem a indica��o da al�quota de ICMS devido no Simples Nacional e o valor do cr�dito correspondente'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 102
      SQL = SQL & ",'Tributada pelo Simples Nacional sem permiss�o de cr�dito'"
      SQL = SQL & ",'classificam-se c�digo as opera��es que n�o permitem a indica��o da al�quota do ICMS devido pelo Simples Nacional e do valor do cr�dito, e n�o estejam abrangidas nas hip�teses dos c�digos 103, 203, 300, 400, 500 e 900'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 103
      SQL = SQL & ",'Isen��o do ICMS no Simples Nacional para faixa de receita bruta'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es praticadas por optantes do Simples Nacional contempladas com isen��o concedida para faixa de receita bruta nos termos da Lei Complementar n. 123 de 2006'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 201
      SQL = SQL & ",'Tributada pelo Simples Nacional com permiss�o de cr�dito e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es  que permitem a indica��o da al�quota do ICMS devido pelo Simples Nacional e do valor cr�dito e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 202
      SQL = SQL & ",'Tributada pelo Simples Nacional sem permiss�o de cr�dito e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es  que n�o permitem a indica��o da al�quota do ICMS devido pelo Simples Nacional e do valor cr�dito, e n�o estejam abrangidas nas hip�teses dos c�digos 103, 203, 300, 400, 500 e 900 e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 203
      SQL = SQL & ",'Isen��o do ICMS no Simples Nacional para a faixa de receita bruta e com cobran�a de ICMS por substitui��o tribut�ria'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es que praticadas por optantes do Simples Nacional contemplados com isen��o para a faixa de receita bruta, mas com ICMS cobrado por substitui��o tribut�ria'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 300
      SQL = SQL & ",'Imune'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es que praticadas por optantes do Simples Nacional contempladas com imunidade do ICMS'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 400
      SQL = SQL & ",'N�o tributada pelo Simples Nacional'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es que praticadas por optantes do Simples NacionaL n�o sujeitas � tributa��o pelo ICMS dentro do Simples Nacional'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 500
      SQL = SQL & ",'ICMS cobrado anteriormente por substitui��o tribut�ria'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es sujeitas exclusivamente ao regime de substitui��o tribut�ria na condi��o de substitu�do tribut�rio ou no caso de antecipa��es'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSON ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 900
      SQL = SQL & ",'Outros'"
      SQL = SQL & ",'classificam-se neste c�digo as opera��es que n�o se enquadrem nos c�digos 101, 102, 103, 201, 202, 203, 300, 400 e 500'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "CST", "") = False Then
      SQL = "CREATE TABLE CST("
      SQL = SQL & " CODIGO NVARCHAR(3) NOT NULL,"
      SQL = SQL & " DESCRICAO NVARCHAR(max) NOT NULL,"
      SQL = SQL & " OBS NVARCHAR(max) NULL,"
      SQL = SQL & " CONSTRAINT PK_TRIBUTACAO_CST PRIMARY KEY CLUSTERED"
      SQL = SQL & " (Codigo Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & "'00'"
      SQL = SQL & ",'Tributada integralmente'"
      SQL = SQL & ",'Tributada integralmente'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 10
      SQL = SQL & ",'Tributada  e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ",'Tributada  e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 20
      SQL = SQL & ",'Com redu��o de base de c�lculo'"
      SQL = SQL & ",'Com redu��o de base de c�lculo'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 30
      SQL = SQL & ",'Isenta ou n�o tributada e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ",'Isenta ou n�o tributada e com cobran�a do ICMS por substitui��o tribut�ria'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 40
      SQL = SQL & ",'Isenta'"
      SQL = SQL & ",'Isenta'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 41
      SQL = SQL & ",'N�o tributada'"
      SQL = SQL & ",'N�o tributada'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 50
      SQL = SQL & ",'Suspens�o'"
      SQL = SQL & ",'Suspens�o'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 51
      SQL = SQL & ",'Diferimento'"
      SQL = SQL & ",'Diferimento'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 60
      SQL = SQL & ",'ICMS cobrado anteriormente por substitui��o tribut�ria'"
      SQL = SQL & ",'ICMS cobrado anteriormente por substitui��o tribut�ria'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 70
      SQL = SQL & ",'Com redu��o de base de c�lculo e cobran�a de ICMS por substitui��o tribut�ria'"
      SQL = SQL & ",'Com redu��o de base de c�lculo e cobran�a de ICMS por substitui��o tribut�ria'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 90
      SQL = SQL & ",'Outras'"
      SQL = SQL & ",'Outras'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "ACESSO", "") = False Then
      SQL = "CREATE TABLE [dbo].[ACESSO]("
      SQL = SQL & " [USUARIO_ID] [int] NOT NULL,"
      SQL = SQL & " [PROGRAMA_ID] Not [Int]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

'=================================
   If ExisteTabela("RETAGUARDA", "PEDIDO", "") = True Then
      If ExisteCampo("RETAGUARDA", "CODG_CLIENTE", "PEDIDO") = True Then
         SQL = "ALTER TABLE PEDIDO ADD CLIENTE_ID BIGINT"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "update PEDIDO set "
         SQL = SQL & " cliente_id = codg_cliente"
   SQL = SQL & " and empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE PEDIDO DROP COLUMN CODG_CLIENTE"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If ExisteTabela("RETAGUARDA", "vwRelVenda", "") = False Then
      SQL = "CREATE VIEW [dbo].[vwRelVenda]"
      SQL = SQL & " AS"
      SQL = SQL & " SELECT EMPRESA.RAZAO_SOCIAL, PEDIDO.NUMR_REQ, PEDIDO.CGCCPF, PEDIDO.CLIENTE_ID, PEDIDO.DT_REQ, PEDIDO.TIPOVENDA_ID, "
      SQL = SQL & " PEDIDO.STATUS, PEDIDO.VALOR_DESCONTO AS Desconto_Cabeca, PEDIDO.NOME_CLIENTE, PEDIDO.VALOR_TOTAL,"
      SQL = SQL & " PEDIDOITEM.CODG_PROD, PEDIDOITEM.QTD_PEDIDA, PEDIDOITEM.VALOR_ITEM, PEDIDOITEM.VALOR_DESCONTO AS Desconto_Item,"
      SQL = SQL & " PRODUTO.DESCRICAO AS Desc_Produto, PRODUTO.PRECO_CUSTO AS Produto_Custo, VENDEDOR.NOME_VEND,"
      SQL = SQL & " TIPOVENDA.DESCRICAO AS Desc_Venda, FORMAPAGTO.DESCRICAO AS Desc_Forma, PEDIDOITEM.PRECO_CUSTO AS Item_Custo"
      SQL = SQL & " FROM EMPRESA INNER JOIN"
      SQL = SQL & " PEDIDO ON EMPRESA.EMPRESA_ID = PEDIDO.EMPRESA_ID INNER JOIN"
      SQL = SQL & " PEDIDOITEM ON PEDIDO.EMPRESA_ID = PEDIDOITEM.EMPRESA_ID AND PEDIDO.NUMR_REQ = PEDIDOITEM.NUMR_REQ INNER JOIN"
      SQL = SQL & " TIPOVENDA ON PEDIDO.EMPRESA_ID = TIPOVENDA.EMPRESA_ID AND PEDIDO.TIPOVENDA_ID = TIPOVENDA.TIPOVENDA_ID INNER JOIN"
      SQL = SQL & " FORMAPAGTO ON TIPOVENDA.formapagto_id = FORMAPAGTO.formapagto_id AND TIPOVENDA.EMPRESA_ID = FORMAPAGTO.EMPRESA_ID INNER JOIN"
      SQL = SQL & " PRODUTO ON PEDIDOITEM.EMPRESA_ID = PRODUTO.EMPRESA_ID AND PEDIDOITEM.CODG_PROD = PRODUTO.CODG_PRODUTO INNER JOIN"
      SQL = SQL & " VENDEDOR ON PEDIDO.VENDEDOR = VENDEDOR.vendedor_id"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   'INICIO EM 11/05/2012
   ' ESTE CAMPO FARA O CONTROLE DE FORMS DESENVOLVIDOS E SOMENTE UTILIZADOS POR AQUELA EMPRESA
   If ExisteCampo("RETAGUARDA", "CGCEMPRESA", "PROGRAMA") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE dbo.PROGRAMA ADD CGCEMPRESA NVARCHAR(20) NULL"

   'FIM  11/05/2012
End Sub

Public Function CONVERTE_DOLAR(Valr_Conversao As Double) As Double
'1 d�lar = 1,84 reais
'x d�lares = 97 reais
'Logo, x = 97/1,84 = 52,72 d�lares

   CONVERTE_DOLAR = 0

   If Valr_Conversao > 0 Then

      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      SQL = "select valor_dolar from CAIXADIA "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and dt_abertura >= '" & Format(Date, "dd/mm/yyyy") & "'"
      SQL = SQL & " and dt_abertura <= '" & Format(Date, "dd/mm/yyyy") & "'"
      SQL = SQL & " and tipo = 'B'"
   SQL = SQL & " and ESTABELECIMENTO_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and NUMERO_CAIXA_CPU = " & NUMERO_CAIXA_CPU

      TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCAIXA.EOF Then _
         If Not IsNull(TabCAIXA.Fields(0).Value) Then _
               If TabCAIXA.Fields(0).Value > 0 Then _
                  CONVERTE_DOLAR = Valr_Conversao / TabCAIXA.Fields(0).Value

      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close
   End If
End Function

Public Function VALR_DOLAR_DIA(DIA_D As Date) As Double
   VALR_DOLAR_DIA = 0

   If IsDate(DIA_D) Then

      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      SQL = "select valor_dolar from CAIXADIA "
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      SQL = SQL & " and dt_abertura >= '" & Format(DIA_D, "dd/mm/yyyy") & "'"
      SQL = SQL & " and dt_abertura <= '" & Format(DIA_D, "dd/mm/yyyy") & "'"
      SQL = SQL & " and tipo = 'B'"
   SQL = SQL & " and ESTABELECIMENTO_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and NUMERO_CAIXA_CPU = " & NUMERO_CAIXA_CPU

      TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCAIXA.EOF Then _
         If Not IsNull(TabCAIXA.Fields(0).Value) Then _
               If TabCAIXA.Fields(0).Value > 0 Then _
                  VALR_DOLAR_DIA = TabCAIXA.Fields(0).Value

      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close
   End If
End Function

Public Sub SP_PESSOA(TIPO_A As String, _
                     ID_N As Long, _
                     CNPJCPF As String, _
                     DESCRICAO As String, _
                     RAZAO As String, _
                     DATA_CAD As String, _
                     SITUACAO As String)

   DESCRICAO = Replace(DESCRICAO, "'", " ")
   RAZAO = Replace(RAZAO, "'", " ")

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   If Trim(TIPO_A) = "C" Then 'consulta
      SQL = "select * from PESSOA "
      SQL = SQL & " where pessoa_id > -1 "

      If ID_N > -1 Then _
         SQL = SQL & " and pessoa_id = " & ID_N

      If Trim(CNPJCPF) <> "" Then _
         SQL = SQL & " and cnpjcpf = '" & CNPJCPF & "'"

      If Trim(DESCRICAO) <> "" Then _
         SQL = SQL & " and DESCRICAO = '" & Trim(DESCRICAO) & "'"

      If Trim(RAZAO) <> "" Then _
         SQL = SQL & " and razao = '" & Trim(RAZAO) & "'"

      'If IsDate(DATA_CAD) And CDate(DATA_CAD) > 0 Then _
         SQL = SQL & " and DATA_CAD = '" & DMA(DATA_CAD) & "'"

      If Trim(SITUACAO) <> "" Then _
         SQL = SQL & " and SITUACAO  = '" & Trim(SITUACAO) & "'"

      TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      Else
         If Trim(TIPO_A) = "" Then
            MsgBox "Informar tipo de instru��o para procedimento."
            Exit Sub
         End If
         If IsNull(ID_N) Then
            MsgBox "Informar ID ou 0 para procedimento."
            Exit Sub
         End If

         If ID_N <= 0 Then
            If Trim(CNPJCPF) = "" Then
               MsgBox "Informar CNPJCPF ."
               Exit Sub
            End If

            If Trim(DESCRICAO) = "" Then
               MsgBox "Informar DESCRICAO/NOME ."
               Exit Sub
            End If

            If Trim(RAZAO) = "" Then
               MsgBox "Informar RAZAO/NOME ."
               Exit Sub
            End If

            'If IsDate(DMA(DATA_CAD)) = "" Then
            '   MsgBox "Informar DATA_CAD ."
            '   Exit Sub
            'End If

            If Trim(SITUACAO) = "" Then
               MsgBox "Informar SITUACAO ."
               Exit Sub
            End If

            ID_N = MAX_ID("pessoa_id", "pessoa", "", "", "", "")

            SQL = "INSERT INTO PESSOA "
            SQL = SQL & " (PESSOA_ID, CnpjCpf, Descricao, RAZAO, DATA_CAD, SITUACAO) "
            SQL = SQL & " VALUES ("
               SQL = SQL & ID_N
               SQL = SQL & ",'" & Trim(CNPJCPF) & "'"
               SQL = SQL & ",'" & Trim(DESCRICAO) & "'"
               SQL = SQL & ",'" & Trim(RAZAO) & "'"
               SQL = SQL & ",'" & DMA(DATA_CAD) & "'"
               SQL = SQL & ",'" & Trim(SITUACAO) & "'"
            SQL = SQL & " )"
            CONECTA_RETAGUARDA.Execute SQL
            Else
               If Trim(CNPJCPF) <> "" Or _
                  Trim(DESCRICAO) <> "" Or _
                  Trim(RAZAO) <> "" Or _
                  IsDate(DATA_CAD) Or _
                  Trim(SITUACAO) <> "" Or _
                  ID_N > -1 _
                  Then

                  SQL = "update PESSOA set "

                  If Trim(CNPJCPF) <> "" Then _
                     SQL = SQL & " CnpjCpf = '" & Trim(CNPJCPF) & "'"

                  If Trim(DESCRICAO) <> "" Then _
                     SQL = SQL & ", Descricao = '" & Trim(DESCRICAO) & "'"

                  If Trim(RAZAO) <> "" Then _
                     SQL = SQL & ", RAZAO = '" & Trim(RAZAO) & "'"

                  If Trim(SITUACAO) <> "" Then _
                     SQL = SQL & ", SITUACAO = '" & Trim(SITUACAO) & "'"

                  SQL = SQL & " where pessoa_id = " & ID_N

                  CONECTA_RETAGUARDA.Execute SQL
               End If
         End If
   End If

   If TabPessoa.State = 1 Then _
      TabPessoa.Close
End Sub

Public Sub GRAVA_RG(CNPJCPF_A As String, NUMR_RG As String, ORGAO_A As String, DT_EXP_D As String)
'On Error GoTo ERRO_TRATA

   If Trim(DT_EXP_D) = "" Then _
      DT_EXP_D = 0

   If PESSOA_ID_N > 0 Then

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from RG "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "INSERT INTO RG "
            SQL = SQL & " (pessoa_id,prop, numero_rg, orgao, dt_exp ) "
         SQL = SQL & " VALUES ("
            SQL = SQL & PESSOA_ID_N
            SQL = SQL & ",'" & Trim(CNPJCPF_A) & "'"
            SQL = SQL & ",'" & Trim(NUMR_RG) & "'"
            SQL = SQL & ",'" & Trim(ORGAO_A) & "'"
            SQL = SQL & "," & DMA(DT_EXP_D)
         SQL = SQL & " )"
         Else
             SQL = "UPDATE RG SET "
             SQL = SQL & " numero_rg = '" & Trim(NUMR_RG) & "'"
             SQL = SQL & ", Orgao = '" & Trim(ORGAO_A) & "'"
             SQL = SQL & ", dt_exp = '" & DMA(DT_EXP_D) & "'"
             SQL = SQL & " Where Prop = '" & Trim(CNPJCPF_A) & "'"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GRAVA_RG"
End Sub

Public Sub GRAVA_FONE(CNPJCPF_A As String, NUMR_FONE As String, DDD_N As Integer, LOCAL_FONE As String, RAMAL_N As String)
'On Error GoTo ERRO_TRATA

   If PESSOA_ID_N > 0 And Trim(NUMR_FONE) <> "" Then
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from FONE "
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numero = '" & Trim(NUMR_FONE) & "'"
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         SQL = "INSERT INTO FONE "
            SQL = SQL & " (PESSOA_ID,PROP,NUMERO,DDD,LOCAL,RAMAL) "
         SQL = SQL & " VALUES ("
            SQL = SQL & PESSOA_ID_N
            SQL = SQL & ",'" & Trim(CNPJCPF_A) & "'"
            SQL = SQL & ",'" & Trim(NUMR_FONE) & "'"
            SQL = SQL & ",0" & Trim(DDD_N)
            SQL = SQL & ",'" & Trim(LOCAL_FONE) & "'"
            SQL = SQL & ",0" & Trim(RAMAL_N)
         SQL = SQL & " )"
         Else
            SQL = "UPDATE FONE SET "
               SQL = SQL & "PROP = '" & Trim(CNPJCPF_A) & "'"
               SQL = SQL & ", NUMERO = '" & Trim(NUMR_FONE) & "'"
               SQL = SQL & ", DDD = 0" & Trim(DDD_N)
               SQL = SQL & ", LOCAL = '" & Trim(LOCAL_FONE) & "'"
               SQL = SQL & ", RAMAL = 0" & Trim(RAMAL_N)
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and numero = '" & Trim(NUMR_FONE) & "'"
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GRAVA_FONE"
End Sub

Public Sub ATUALIZA_TABELA_FAMILIAPRODUTO()
'On Error GoTo ERRO_TRATA

   If ExisteTabela("RETAGUARDA", "FAMILIAPRODUTO", "") = False Then
      SQL = "CREATE TABLE [dbo].[FAMILIAPRODUTO]("
      SQL = SQL & " [FAMILIAPRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CODG_FAMILIA] [nvarchar](10) NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](60) NOT NULL,"
      SQL = SQL & " [UNIDADE_MEDIDA] [nvarchar](4) NULL,"
      SQL = SQL & " [DESC_UNIDADE_MEDIDA] [nvarchar](50) NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If ExisteCampo("RETAGUARDA", "FAMILIAPRODUTO_ID", "FAMILIAPRODUTO") = True Then
            Altera��o_Defini��o_Campo_Tabela "FAMILIAPRODUTO_ID", "BIGINT NOT NULL", "FAMILIAPRODUTO"
            Else: CONECTA_RETAGUARDA.Execute "ALTER TABLE FAMILIAPRODUTO ADD FAMILIAPRODUTO_ID BIGINT NOT NULL"
         End If
   End If

   If ExisteTabela("RETAGUARDA", "pk_FAMILIAPRODUTO", "") = False Then
      SQL = "ALTER TABLE FAMILIAPRODUTO ADD CONSTRAINT pk_FAMILIAPRODUTO PRIMARY KEY (FAMILIAPRODUTO_ID)"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "ATUALIZA_TABELA_FAMILIAPRODUTO"
End Sub

Public Sub CHECA_TABELA_PRODUTO()
'On Error GoTo ERRO_TRATA

   ATUALIZA_TABELA_FAMILIAPRODUTO

   If ExisteCampo("RETAGUARDA", "nacional", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.PRODUTO.nacional'" & "," & "'ORIGEM_MERCADO'" & "," & "'COLUMN'"

   If ExisteCampo("RETAGUARDA", "codg_fornec", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.PRODUTO.codg_fornec'" & "," & "'FORNECEDOR_ID'" & "," & "'COLUMN'"
   
   If ExisteCampo("RETAGUARDA", "QTDE", "PRODUTO") = True Then _
      Altera��o_Defini��o_Campo_Tabela "QTDE", "NUMERIC(18,3)", "PRODUTO"

   If ExisteCampo("RETAGUARDA", "qtd_BALCAO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.PRODUTO.qtd_BALCAO'" & "," & "'QTDE_RETIDO'" & "," & "'COLUMN'"

   If ExisteCampo("RETAGUARDA", "qtde_retido", "PRODUTO") = True Then _
      Altera��o_Defini��o_Campo_Tabela "qtde_retido", "NUMERIC(18,3)", "PRODUTO"

   If ExisteCampo("RETAGUARDA", "CODG_PRODUTO", "PRODUTO") = True Then _
      Altera��o_Defini��o_Campo_Tabela "CODG_PRODUTO", "NVARCHAR(100)", "PRODUTO"

   'If ExisteTabela("RETAGUARDA", "pk_PRODUTO","") = True Then _
      CONECTA_RETAGUARDA.Execute SQL = "ALTER TABLE PRODUTO DROP CONSTRAINT pk_PRODUTO "

   If ExisteCampo("RETAGUARDA", "PRODUTO_ID", "PRODUTO") = True Then
      Altera��o_Defini��o_Campo_Tabela "PRODUTO_ID", "BIGINT", "PRODUTO"
      Else: CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PRODUTO_ID BIGINT"
   End If

   If ExisteCampo("RETAGUARDA", "PESO_LIQUIDO", "PRODUTO") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PESO_LIQUIDO FLOAT"
      Else: Altera��o_Defini��o_Campo_Tabela "PESO_LIQUIDO", "FLOAT", "PRODUTO"
   End If

   If ExisteCampo("RETAGUARDA", "PESO_BRUTO", "PRODUTO") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PESO_BRUTO NUMERIC(18,3)"
      Else: Altera��o_Defini��o_Campo_Tabela "PESO_BRUTO", "NUMERIC(18,3)", "PRODUTO"
   End If

   If ExisteCampo("RETAGUARDA", "FAMILIAPRODUTO_ID", "PRODUTO") = True Then
      Altera��o_Defini��o_Campo_Tabela "FAMILIAPRODUTO_ID", "BIGINT", "PRODUTO"
      Else: CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD FAMILIAPRODUTO_ID BIGINT"
   End If

   If ExisteCampo("RETAGUARDA", "REFERENCIA", "PRODUTO") = True Then _
      Altera��o_Defini��o_Campo_Tabela "REFERENCIA", "NVARCHAR(200)", "PRODUTO"

   If ExisteCampo("RETAGUARDA", "CODG_USUARIO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.PRODUTO.CODG_USUARIO'" & "," & "'USUARIO_ID'" & "," & "'COLUMN'"

   If ExisteCampo("RETAGUARDA", "TAMANHO", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD TAMANHO BIGINT NULL"

   If ExisteCampo("RETAGUARDA", "MARCA_ID", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD MARCA_ID BIGINT NULL"

   If ExisteCampo("RETAGUARDA", "PRODUTO_BALANCA", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PRODUTO_BALANCA bit"

   If ExisteCampo("RETAGUARDA", "CFOP", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN CFOP"

   If ExisteCampo("RETAGUARDA", "CODIGO_ORIGINAL", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN CODIGO_ORIGINAL"

   If ExisteCampo("RETAGUARDA", "ALIQUOTA_SUBST_TRIBUTARIA", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN ALIQUOTA_SUBST_TRIBUTARIA"

   If ExisteCampo("RETAGUARDA", "VALOR_VENDA", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN VALOR_VENDA"

   If ExisteCampo("RETAGUARDA", "VALOR_CUSTO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN VALOR_CUSTO"

   If ExisteCampo("RETAGUARDA", "GRUPO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN GRUPO"

   If ExisteCampo("RETAGUARDA", "NATUREZA", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN NATUREZA"

   If ExisteCampo("RETAGUARDA", "CLASSIFICACAOFISCAL", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN CLASSIFICACAOFISCAL"

   If ExisteCampo("RETAGUARDA", "IMAGEM", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN IMAGEM"

   If ExisteCampo("RETAGUARDA", "STATUS", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN STATUS"

   If ExisteTabela("RETAGUARDA", "pk_PRODUTO", "") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD CONSTRAINT pk_PRODUTO PRIMARY KEY (PRODUTO_ID)"

   If ExisteTabela("RETAGUARDA", "FK_PRODUTO_EMPRESA", "") = False Then
      SQL = "ALTER TABLE [dbo].[PRODUTO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
      SQL = SQL & " References [dbo].[Empresa]([EMPRESA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTO] CHECK CONSTRAINT [FK_PRODUTO_EMPRESA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_PRODUTO_FAMILIAPRODUTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[PRODUTO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTO_FAMILIAPRODUTO] FOREIGN KEY([FAMILIAPRODUTO_ID])"
      SQL = SQL & " References [dbo].[FAMILIAPRODUTO]([FAMILIAPRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTO] CHECK CONSTRAINT [FK_PRODUTO_FAMILIAPRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   SQL = "update PEDIDOITEM set tipo_reg = 'PC' "
   SQL = SQL & " where tipo_reg Is Null "
   CONECTA_RETAGUARDA.Execute SQL

'If INDR_ESTQ_NEGATIVO = False Then
   SQL = "update produto set qtde_retido = 0"
   SQL = SQL & " where qtde_retido < 0 "
   CONECTA_RETAGUARDA.Execute SQL
'End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "MDLGERAL", "CHECA_TABELA_PRODUTO"
End Sub

Public Sub CHECA_TABELA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   If ExisteTabela("RETAGUARDA", "ESTOQUE", "") = False Then
      SQL = "CREATE TABLE [dbo].[ESTOQUE]("
      SQL = SQL & " [ESTOQUE_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [QTDE_ESTOQUE] [FLOAT] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_ESTOQUE] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([ESTOQUE_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],"
      SQL = SQL & " CONSTRAINT [IX_ESTOQUE_PRODUTO] UNIQUE NONCLUSTERED"
      SQL = SQL & " ([ESTABELECIMENTO_ID] ASC,[PRODUTO_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, "
      SQL = SQL & " ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_ESTOQUE_ESTABELECIMENTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[ESTOQUE] "
      SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_ESTOQUE_ESTABELECIMENTO] "
      SQL = SQL & " FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTOQUE] CHECK CONSTRAINT [FK_ESTOQUE_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_ESTOQUE_PRODUTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[ESTOQUE] "
      SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_ESTOQUE_PRODUTO] "
      SQL = SQL & " FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTOQUE] CHECK CONSTRAINT [FK_ESTOQUE_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "CHECA_TABELA_ESTOQUE"
End Sub

Public Sub MATA_TABELAS()
On Error Resume Next

   If ExisteTabela("RETAGUARDA", "DEBITO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE DEBITO"

   If ExisteTabela("RETAGUARDA", "ListaPrecos", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ListaPrecos"

   If ExisteTabela("RETAGUARDA", "NATUREZAOPERACAOPRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE NATUREZAOPERACAOPRODUTO"

   If ExisteTabela("RETAGUARDA", "CABCOMPR", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE CABCOMPR"

   If ExisteTabela("RETAGUARDA", "CABDEVENT", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE CABDEVENT"

   If ExisteTabela("RETAGUARDA", "CABDEVSAI", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE CABDEVSAI"

   If ExisteTabela("RETAGUARDA", "BancoNaturezaOcorrencia", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE BancoNaturezaOcorrencia"

   If ExisteTabela("RETAGUARDA", "BancoParametroRetorno", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE BancoParametroRetorno"

   If ExisteTabela("RETAGUARDA", "ALTERACAO_CLIENTE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ALTERACAO_CLIENTE"

   If ExisteTabela("RETAGUARDA", "ALTERACAO_PRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ALTERACAO_PRODUTO"

   If ExisteTabela("RETAGUARDA", "ITEMRECEBIMENTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ITEMRECEBIMENTO"

   If ExisteTabela("RETAGUARDA", "ITENSCOM", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ITENSCOM"

   If ExisteTabela("RETAGUARDA", "altera��o_CLIENTE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE altera��o_CLIENTE"

   If ExisteTabela("RETAGUARDA", "ALIQUOTA_UF", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ALIQUOTA_UF"

   If ExisteTabela("RETAGUARDA", "BANCOFINA", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE BANCOFINA"

   If ExisteTabela("RETAGUARDA", "altera��o_PRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE altera��o_PRODUTO"

   If ExisteTabela("RETAGUARDA", "vwRel_EstoqueVenda", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE vwRel_EstoqueVenda"

   If ExisteTabela("RETAGUARDA", "consulta_tributa��o", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view consulta_tributa��o"

   If ExisteTabela("RETAGUARDA", "CADASTRO_IVA_UF_PRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP table CADASTRO_IVA_UF_PRODUTO"

   If ExisteTabela("RETAGUARDA", "consulta familia", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view consulta familia"

   If ExisteTabela("RETAGUARDA", "consulta_venda", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view consulta_venda"

   If ExisteTabela("RETAGUARDA", "familia_consulta", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view familia_consulta"

   If ExisteTabela("RETAGUARDA", "EMPRESAPARAMETRO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE EMPRESAPARAMETRO"

   If ExisteTabela("RETAGUARDA", "DEVOLUCAO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE DEVOLUCAO"

   If ExisteTabela("RETAGUARDA", "DEVOLUCAOTEMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE DEVOLUCAOTEMP"

   If ExisteTabela("RETAGUARDA", "TIPO60MESTRE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60MESTRE"

   If ExisteTabela("RETAGUARDA", "TIPO60ANALITICO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60ANALITICO"

   If ExisteTabela("RETAGUARDA", "TIPO60ITEM", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60ITEM"

   If ExisteTabela("RETAGUARDA", "TIPO60RESUMODIA", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60RESUMODIA"

   If ExisteTabela("RETAGUARDA", "TIPO60RESUMOMENSAL", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60RESUMOMENSAL"

   If ExisteTabela("RETAGUARDA", "NOTATEMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE NOTATEMP"

   If ExisteTabela("RETAGUARDA", "FUNCIONARIO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE FUNCIONARIO"

   If ExisteTabela("RETAGUARDA", "FUNCIONARIOCONVENIO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE FUNCIONARIOCONVENIO"

   If ExisteTabela("RETAGUARDA", "GRUPOPRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE GRUPOPRODUTO"

   If ExisteTabela("RETAGUARDA", "IMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE IMP"

   If ExisteTabela("RETAGUARDA", "PEDIDOTEMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE PEDIDOTEMP"

   If ExisteTabela("RETAGUARDA", "SERIEIMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE SERIEIMP"
End Sub

Public Sub ATUALIZA_TABELA_PESSOA()
   If ExisteTabela("RETAGUARDA", "PESSOA", "") = False Then
      SQL = "CREATE TABLE [dbo].[PESSOA]("
      SQL = SQL & " [PESSOA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CNPJCPF] [NVARCHAR](14) NOT NULL,"
      SQL = SQL & " [DESCRICAO] [NVARCHAR](max) NOT NULL,"
      SQL = SQL & " [RAZAO] [NVARCHAR](max) NOT NULL,"
      SQL = SQL & " [DATA_CAD] [datetime] NOT NULL,"
      SQL = SQL & " [SITUACAO] [NVARCHAR](1) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_PESSOA] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([PESSOA_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If ExisteTabela("RETAGUARDA", "pk_PESSOA", "") = False Then
            SQL = "ALTER TABLE PESSOA ADD CONSTRAINT pk_PESSOA PRIMARY KEY (PESSOA_ID)"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If ExisteCampo("RETAGUARDA", "TIPO_REG", "PESSOA") = True Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE PESSOA DROP COLUMN TIPO_REG"
   End If
   'CREATE INDEX IX_CNPJCPF ON PESSOA(CNPJCPF) WITH (ONLINE=ON, SORT_IN_TEMPDB=ON)
End Sub

Public Sub ATUALIZA_TABELA_EMPRESA()
   If ExisteTabela("RETAGUARDA", "EMPRESA", "") = True Then
      If ExisteCampo("RETAGUARDA", "INDR_INDUSTRIA", "EMPRESA") = True Then _
         Altera��o_Defini��o_Campo_Tabela "INDR_INDUSTRIA", "BIT", "EMPRESA"

      If ExisteCampo("RETAGUARDA", "PESSOA_ID", "EMPRESA") = True Then _
         Altera��o_Defini��o_Campo_Tabela "PESSOA_ID", "BIGINT NOT NULL", "EMPRESA"

      If ExisteCampo("RETAGUARDA", "Nome_imp", "EMPRESA") = True Then
         SQL = "EXEC sp_rename " & "'dbo.EMPRESA.Nome_imp'" & "," & "'IMPRESSORA_FISCAL'" & "," & "'COLUMN'"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      Altera��o_Defini��o_Campo_Tabela "IMPRESSORA_FISCAL", "INT", "EMPRESA"

      If ExisteCampo("RETAGUARDA", "FONE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN FONE"

      If ExisteCampo("RETAGUARDA", "SEQ_CONSULTA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_CONSULTA"

      If ExisteCampo("RETAGUARDA", "SEQ_FUNCIONARIO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_FUNCIONARIO"

      If ExisteCampo("RETAGUARDA", "ECF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ECF"

      If ExisteCampo("RETAGUARDA", "NUMR_PEDIDO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN NUMR_PEDIDO"

      If ExisteCampo("RETAGUARDA", "SEQ_NOTA_ENTRADA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_NOTA_ENTRADA"

      If ExisteCampo("RETAGUARDA", "SEQ_PEDIDO_ENTRADA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_PEDIDO_ENTRADA"

      If ExisteCampo("RETAGUARDA", "CONTROLA_ESTOQUE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN CONTROLA_ESTOQUE"

      If ExisteCampo("RETAGUARDA", "ALIQUOTA_DENTRO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ALIQUOTA_DENTRO"

      If ExisteCampo("RETAGUARDA", "ALIQUOTA_FORA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ALIQUOTA_FORA"

      If ExisteCampo("RETAGUARDA", "PERC_ICMS_ENTRADA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PERC_ICMS_ENTRADA"

      If ExisteCampo("RETAGUARDA", "SENHA_BANCO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SENHA_BANCO"

      If ExisteCampo("RETAGUARDA", "USUARIO_BANCO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN USUARIO_BANCO"

      If ExisteCampo("RETAGUARDA", "CODIGOBARRAPESOQTD", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN CODIGOBARRAPESOQTD"

      If ExisteCampo("RETAGUARDA", "CAMINHOREL", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN CAMINHOREL"

      If ExisteCampo("RETAGUARDA", "PERC_DESCONTO_CONV", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PERC_DESCONTO_CONV"

      If ExisteCampo("RETAGUARDA", "PSPermiteAlterarValor", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSPermiteAlterarValor"

      If ExisteCampo("RETAGUARDA", "Tipo_Assinatura", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN Tipo_Assinatura"

      If ExisteCampo("RETAGUARDA", "VersaoPaf", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN VersaoPaf"

      If ExisteCampo("RETAGUARDA", "PCNumeroECF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCNumeroECF"

      If ExisteCampo("RETAGUARDA", "desabilitareducaoz", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN desabilitareducaoz"

      If ExisteCampo("RETAGUARDA", "PCUsaBalancaCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCUsaBalancaCaixa"

      If ExisteCampo("RETAGUARDA", "PSAntesIniciarVendaCartao", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAntesIniciarVendaCartao"

      If ExisteCampo("RETAGUARDA", "PCUsaLeitorSerial", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCUsaLeitorSerial"

      If ExisteCampo("RETAGUARDA", "PCImpressoraUsaPorta", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCImpressoraUsaPorta"

      If ExisteCampo("RETAGUARDA", "PCPortaBalancaCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCPortaBalancaCaixa"

      If ExisteCampo("RETAGUARDA", "PCPorta", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCPorta"

      If ExisteCampo("RETAGUARDA", "PCPortaLeitorSerial", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCPortaLeitorSerial"

      If ExisteCampo("RETAGUARDA", "PSAliquotas", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAliquotas"

      If ExisteCampo("RETAGUARDA", "PBDBancoDadosCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PBDBancoDadosCaixa"

      If ExisteCampo("RETAGUARDA", "PSSistemaControlaCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSSistemaControlaCaixa"

      If ExisteCampo("RETAGUARDA", "PSAntesIniciarVendaCPF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAntesIniciarVendaCPF"

      If ExisteCampo("RETAGUARDA", "PTBandeiraTecBan", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraTecBan"

      If ExisteCampo("RETAGUARDA", "PTBandeiraVisa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraVisa"

      If ExisteCampo("RETAGUARDA", "PTBandeiraMasterCard", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraMasterCard"

      If ExisteCampo("RETAGUARDA", "PTBandeiraAmericanExpress", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraAmericanExpress"

      If ExisteCampo("RETAGUARDA", "PSPermiteDesconto", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSPermiteDesconto"

      If ExisteCampo("RETAGUARDA", "PSAbrirGaveta", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAbrirGaveta"

      If ExisteCampo("RETAGUARDA", "PTUsaTef", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTUsaTef"

      If ExisteCampo("RETAGUARDA", "ImprimeOrcNormalFiscal", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ImprimeOrcNormalFiscal"

      If ExisteCampo("RETAGUARDA", "VERSAO_CNIECF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN VERSAO_CNIECF"

      If ExisteCampo("RETAGUARDA", "Aliquota", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN Aliquota"

      If ExisteCampo("RETAGUARDA", "TipoTefDiscadoDedicado", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN TipoTefDiscadoDedicado"

      If ExisteCampo("RETAGUARDA", "Empresa_Regime_TARE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN Empresa_Regime_TARE"

      If ExisteCampo("RETAGUARDA", "super_simples", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN super_simples"

      If ExisteCampo("RETAGUARDA", "OPTANTE_TARE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN OPTANTE_TARE"

      If ExisteCampo("RETAGUARDA", "le_deposito", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD LE_DEPOSITO BIT "

      If ExisteCampo("RETAGUARDA", "PAR", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD PAR NVARCHAR(8) "

'============================================================ DROP MIGRA ESTABELECIMENTO
If ExisteCampo("RETAGUARDA", "IMPRESSORA_FISCAL", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN IMPRESSORA_FISCAL"

If ExisteCampo("RETAGUARDA", "usa_impfiscal", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN usa_impfiscal"

If ExisteCampo("RETAGUARDA", "indr_industria", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN indr_industria"

If ExisteCampo("RETAGUARDA", "LIBERA_DESCONTO", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN LIBERA_DESCONTO"

If ExisteCampo("RETAGUARDA", "usa_nfe", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN usa_nfe"

If ExisteCampo("RETAGUARDA", "RECEBE_PEDIDO_VENDA", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN RECEBE_PEDIDO_VENDA"

If ExisteCampo("RETAGUARDA", "ESTOQUE_NEGATIVO", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ESTOQUE_NEGATIVO"

If ExisteCampo("RETAGUARDA", "USA_TEF", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN USA_TEF"

If ExisteCampo("RETAGUARDA", "LEI_12741", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN LEI_12741"

If ExisteCampo("RETAGUARDA", "UsaCobBancaria", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN UsaCobBancaria"

If ExisteCampo("RETAGUARDA", "INSTRUCAO_BOLETO", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN INSTRUCAO_BOLETO"

If ExisteCampo("RETAGUARDA", "INSTRUCAO_FISCO", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN INSTRUCAO_FISCO"

If ExisteCampo("RETAGUARDA", "BAIXA_ESTOQUE_REQ", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN BAIXA_ESTOQUE_REQ"

If ExisteCampo("RETAGUARDA", "PERC_JUROS_ATRAZO", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PERC_JUROS_ATRAZO"

If ExisteCampo("RETAGUARDA", "QTD_DIAS_ATRAZO", "EMPRESA") = True Then _
   CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN QTD_DIAS_ATRAZO"

'============================================================

CONT_N = 0

      If ExisteCampo("RETAGUARDA", "cgc", "EMPRESA") = True Then
         If TabEmpresa.State = 1 Then _
            TabEmpresa.Close

         SQL = "select empresa.* from EMPRESA"

   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

         TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabEmpresa.EOF
            CONT_N = CONT_N + 1
            CNPJCPF_A = "" & TabEmpresa.Fields("cgc").Value
            NOME_A = "" & TabEmpresa.Fields("nome_fant").Value
            RAZAO_A = "" & TabEmpresa.Fields("razao_social").Value
            DT_EXP_D = Date

            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            SQL = "select pessoa_id from PESSOA"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
            TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabPessoa.EOF Then
               If TabPessoa.State = 1 Then _
                  TabPessoa.Close

               SP_PESSOA "I", _
                         0, _
                         Trim(cnpfcpf_a), _
                         Trim(NOME_A), _
                         Trim(RAZAO_A), _
                         DT_EXP_D, _
                         Trim("A")

               If TabPessoa.State = 1 Then _
                  TabPessoa.Close

               SQL = "select pessoa_id from PESSOA"
               SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
               TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabPessoa.EOF Then
                  SQL = "update EMPRESA set pessoa_id = " & TabPessoa.Fields(0).Value

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

                  SQL = SQL & " and CGC = '" & Trim(CNPJCPF_A) & "'"
                  CONECTA_RETAGUARDA.Execute SQL
               End If
            End If
            If TabPessoa.State = 1 Then _
               TabPessoa.Close

            TabEmpresa.MoveNext
         Wend
         If TabEmpresa.State = 1 Then _
            TabEmpresa.Close
      End If

      If ExisteCampo("RETAGUARDA", "SEQ_REQORC", "EMPRESA") = True Then
         SQL = "EXEC sp_rename " & "'dbo.EMPRESA.SEQ_REQORC'" & "," & "'SEQ_PEDIDO'" & "," & "'COLUMN'"
         CONECTA_RETAGUARDA.Execute SQL

         Altera��o_Defini��o_Campo_Tabela "SEQ_PEDIDO", "BIGINT", "EMPRESA"
      End If

      If ExisteTabela("RETAGUARDA", "pk_EMPRESA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD CONSTRAINT pk_EMPRESA PRIMARY KEY (EMPRESA_ID)"
   
      If ExisteTabela("RETAGUARDA", "FK_EMPRESA_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[EMPRESA]  WITH CHECK ADD  CONSTRAINT [FK_EMPRESA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[EMPRESA] CHECK CONSTRAINT [FK_EMPRESA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
End Sub

Public Sub ATUALIZA_TABELA_TRANSPORTADORA()
   If ExisteTabela("RETAGUARDA", "TRANSPORTADORA", "") = True Then
      If ExisteCampo("RETAGUARDA", "PESSOA_ID", "TRANSPORTADORA") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA ADD PESSOA_ID BIGINT"
         Else: Altera��o_Defini��o_Campo_Tabela "PESSOA_ID", "BIGINT NOT NULL", "TRANSPORTADORA"
      End If

      If ExisteCampo("RETAGUARDA", "TRANSP_ID", "TRANSPORTADORA") = True Then _
         Altera��o_Defini��o_Campo_Tabela "TRANSP_ID", "BIGINT NOT NULL", "TRANSPORTADORA"

      If ExisteTabela("RETAGUARDA", "pk_TRANSPORTADORA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TRANSPORTADORA ADD CONSTRAINT pk_TRANSPORTADORA PRIMARY KEY (TRANSP_ID)"

      If ExisteTabela("RETAGUARDA", "FK_TRANSPORTADORA_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA]  WITH CHECK ADD  CONSTRAINT [FK_TRANSPORTADORA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA] CHECK CONSTRAINT [FK_TRANSPORTADORA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_TRANSPORTADORA_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA]  WITH CHECK ADD  CONSTRAINT [FK_TRANSPORTADORA_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[TRANSPORTADORA] CHECK CONSTRAINT [FK_TRANSPORTADORA_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
End Sub

Public Sub ATUALIZA_TABELA_USUARIO()
   If ExisteTabela("RETAGUARDA", "USUARIO", "") = True Then
      If ExisteCampo("RETAGUARDA", "PESSOA_ID", "USUARIO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO ADD PESSOA_ID BIGINT"
         Else: Altera��o_Defini��o_Campo_Tabela "PESSOA_ID", "BIGINT", "USUARIO"
      End If

      If ExisteCampo("RETAGUARDA", "USUARIO_ID", "USUARIO") = True Then _
         Altera��o_Defini��o_Campo_Tabela "USUARIO_ID", "BIGINT NOT NULL", "USUARIO"


      If ExisteCampo("RETAGUARDA", "PERC_DESCONTO", "USUARIO") = True Then _
         Altera��o_Defini��o_Campo_Tabela "PERC_DESCONTO", "FLOAT", "USUARIO"

      If ExisteCampo("RETAGUARDA", "PERC_COMISSAO", "USUARIO") = True Then _
         Altera��o_Defini��o_Campo_Tabela "PERC_COMISSAO", "FLOAT", "USUARIO"


      If ExisteTabela("RETAGUARDA", "pk_USUARIO", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO ADD CONSTRAINT pk_USUARIO PRIMARY KEY (USUARIO_ID)"

      If ExisteTabela("RETAGUARDA", "FK_USUARIO_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[USUARIO]  WITH CHECK ADD  CONSTRAINT [FK_USUARIO_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[USUARIO] CHECK CONSTRAINT [FK_USUARIO_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_USUARIO_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[USUARIO]  WITH CHECK ADD  CONSTRAINT [FK_USUARIO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[USUARIO] CHECK CONSTRAINT [FK_USUARIO_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
End Sub

Public Sub LIBERA_DESCONTO()
TRAVEIS_DESCONTO:

   PERC_DESCONTO_USUARIO = 0
   VALOR_TOTAL_DESCONTO_N = 0
   PERC_DESCONTO_N = 0
   CRITERIO = ""

   frmVENDADESCONTO.Show 1

   If TabUSU.State = 1 Then _
      TabUSU.Close

   SQL = "select perc_desconto from USUARIO "
   SQL = SQL & " where usuario_id = " & CODG_USU_N
   TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabUSU.EOF Then _
      If Not IsNull(TabUSU.Fields(0).Value) Then _
         PERC_DESCONTO_USUARIO = TabUSU.Fields(0).Value
   If TabUSU.State = 1 Then _
      TabUSU.Close

   If PERC_DESCONTO_N > PERC_DESCONTO_USUARIO Then
      Msg = "Limite de desconto ultrapassado, deseja liberar com senha superior ?"
      PERGUNTA Msg, vbYesNo + 32, "Desconto Pedido Venda", "DEMO.HLP", 1000
      If RESPOSTA = vbYes Then
         frmSenha.Show 1

         SQL = "select * from USUARIO "
         SQL = SQL & " where senha = '" & Trim(CRITERIO) & "'"
         TabUSU.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabUSU.EOF Then

            If IsNull(TabUSU.Fields("tipo").Value) Then
               MsgBox "N�o permitido, tipo usu�rio n�o informado."
               GoTo TRAVEIS_DESCONTO
            End If

            If TabUSU.Fields("tipo").Value >= 4 And TabUSU.Fields("tipo").Value <= 5 Then
               Else
                  MsgBox "N�o permitido, faixa de desconto n�o cadastrada para este usu�rio."
                  GoTo TRAVEIS_DESCONTO
            End If

            USU_LIBERA_VENDA_N = TabUSU.Fields("usuario_id").Value
            Else
               MsgBox "N�o permitido."
               GoTo TRAVEIS_DESCONTO
         End If

         If TabUSU.State = 1 Then _
            TabUSU.Close
         Else
            VALOR_TOTAL_DESCONTO_N = 0
            PERC_DESCONTO_N = 0
            GoTo TRAVEIS_DESCONTO
      End If
   End If

   CRITERIO = ""
End Sub

Public Sub CHECA_TABELA_ESTABELECIMENTO()
   If ExisteTabela("RETAGUARDA", "ESTABELECIMENTO", "") = True Then

      If ExisteTabela("RETAGUARDA", "IX_ESTABELECIMENTO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO DROP CONSTRAINT IX_ESTABELECIMENTO"

      If ExisteCampo("RETAGUARDA", "INTEGRA", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO drop column INTEGRA"

      If ExisteCampo("RETAGUARDA", "CODG_FOLHA", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO DROP COLUMN CODG_FOLHA"

      If ExisteCampo("RETAGUARDA", "servidor", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO drop column servidor"

      If ExisteCampo("RETAGUARDA", "nome_banco", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO drop column nome_banco"

      If ExisteCampo("RETAGUARDA", "ESTABELECIMENTO_ID", "ESTABELECIMENTO") = True Then _
         Altera��o_Defini��o_Campo_Tabela "ESTABELECIMENTO_ID", "BIGINT NOT NULL", "ESTABELECIMENTO"

      If ExisteCampo("RETAGUARDA", "LOCALIZACAO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LOCALIZACAO NVARCHAR(MAX)"

      If ExisteCampo("RETAGUARDA", "CONTROLE_ESTOQUE", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CONTROLE_ESTOQUE BIT"

'==========================================================================================================
'1=DICADO ; 2=IP ; 3=DEDICADO
      If ExisteCampo("RETAGUARDA", "TIPO_TEF", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD TIPO_TEF INT"

'IMPRESSORA_FISCAL
      If ExisteCampo("RETAGUARDA", "IMPRESSORA_FISCAL", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD IMPRESSORA_FISCAL int "

'USA_IMPFISCAL
      If ExisteCampo("RETAGUARDA", "USA_IMPFISCAL", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD USA_IMPFISCAL bit "

'INDR_INDUSTRIA
      If ExisteCampo("RETAGUARDA", "INDR_INDUSTRIA", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD INDR_INDUSTRIA bit "

'LIBERA_DESCONTO
      If ExisteCampo("RETAGUARDA", "LIBERA_DESCONTO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LIBERA_DESCONTO bit "

'USA_NFe
      If ExisteCampo("RETAGUARDA", "USA_NFe", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD USA_NFe bit "

'RECEBE_PEDIDO_VENDA
      If ExisteCampo("RETAGUARDA", "RECEBE_PEDIDO_VENDA", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD RECEBE_PEDIDO_VENDA bit "

'ESTOQUE_NEGATIVO
      If ExisteCampo("RETAGUARDA", "ESTOQUE_NEGATIVO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD ESTOQUE_NEGATIVO bit "

'USA_TEF
      If ExisteCampo("RETAGUARDA", "USA_TEF", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD USA_TEF bit "

'LEI_12741
      If ExisteCampo("RETAGUARDA", "LEI_12741", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LEI_12741 bit "

'INSTRUCAO_BOLETO
      If ExisteCampo("RETAGUARDA", "INSTRUCAO_BOLETO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD INSTRUCAO_BOLETO NVARCHAR(MAX)"

'BAIXA_ESTOQUE_REQ
      If ExisteCampo("RETAGUARDA", "BAIXA_ESTOQUE_REQ", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD BAIXA_ESTOQUE_REQ BIT "

'PERC_JUROS_ATRAZO
      If ExisteCampo("RETAGUARDA", "PERC_JUROS_ATRAZO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD PERC_JUROS_ATRAZO FLOAT "

'QTD_DIAS_ATRAZO
      If ExisteCampo("RETAGUARDA", "QTD_DIAS_ATRAZO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD QTD_DIAS_ATRAZO INT"

'==========================================================================================================

      If ExisteTabela("RETAGUARDA", "pk_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE ESTABELECIMENTO ADD CONSTRAINT pk_ESTABELECIMENTO PRIMARY KEY (ESTABELECIMENTO_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_ESTABELECIMENTO_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[ESTABELECIMENTO]  WITH CHECK ADD  CONSTRAINT [FK_ESTABELECIMENTO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[Empresa]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[ESTABELECIMENTO] CHECK CONSTRAINT [FK_ESTABELECIMENTO_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

   End If
End Sub

Public Sub ATUALIZA_TABELA_ENTREGA()
   If ExisteTabela("RETAGUARDA", "ENTREGA", "") = False Then
      SQL = "CREATE TABLE [dbo].[ENTREGA]("
      SQL = SQL & " [ENTREGA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PESSOA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [EMPRESA_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
      SQL = SQL & " [DT_AGENDA] [datetime] NULL,"
      SQL = SQL & " [DT_ENTREGA] [datetime] NULL,"
      SQL = SQL & " [CEP] [NVARCHAR](8) NULL,"
      SQL = SQL & " [RUA] [NVARCHAR](50) NULL,"
      SQL = SQL & " [COMPLEMENTO] [NVARCHAR](50) NULL,"
      SQL = SQL & " [BAIRRO] [NVARCHAR](50) NULL,"
      SQL = SQL & " [CIDADE] [NVARCHAR](50) NULL,"
      SQL = SQL & " [UF] [NVARCHAR](2) NULL,"
      SQL = SQL & " [ENTREGADOR_ID] [bigint] ,"
      SQL = SQL & " [MONTADOR_ID] [bigint] ,"

      SQL = SQL & " [ENTREGADOR] [NVARCHAR](30) ,"
      SQL = SQL & " [MONTADOR] [NVARCHAR](30) ,"
      
      SQL = SQL & " CONSTRAINT [pk_ENTREGA] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([ENTREGA_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
      SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_PEDIDO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_PESSOA] FOREIGN KEY([PESSOA_ID])"
      SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_PESSOA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "ENTREGA", "") = True Then
      If ExisteCampo("RETAGUARDA", "ENTREGA_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ENTREGA_ID BIGINT NOT NULL"

      If ExisteCampo("RETAGUARDA", "PESSOA_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD PESSOA_ID BIGINT NOT NULL"

      If ExisteCampo("RETAGUARDA", "PEDIDO_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD PEDIDO_ID BIGINT NOT NULL"

      If ExisteCampo("RETAGUARDA", "DT_CAD", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD DT_CAD DATETIME NOT NULL"

      If ExisteCampo("RETAGUARDA", "DT_AGENDA", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD DT_AGENDA DATETIME "
   
      If ExisteCampo("RETAGUARDA", "DT_ENTREGA", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD DT_ENTREGA DATETIME "

      If ExisteCampo("RETAGUARDA", "CEP", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD CEP NVARCHAR(8)"
   
      If ExisteCampo("RETAGUARDA", "RUA", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD RUA NVARCHAR(50)"

      If ExisteCampo("RETAGUARDA", "COMPLEMENTO", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD COMPLEMENTO NVARCHAR(50)"

      If ExisteCampo("RETAGUARDA", "BAIRRO", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD BAIRRO NVARCHAR(50)"

      If ExisteCampo("RETAGUARDA", "CIDADE", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD CIDADE NVARCHAR(50)"

      If ExisteCampo("RETAGUARDA", "UF", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD UF NVARCHAR(2)"

      If ExisteCampo("RETAGUARDA", "ENTREGADOR_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ENTREGADOR_ID BIGINT "

      If ExisteCampo("RETAGUARDA", "MONTADOR_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD MONTADOR_ID BIGINT "

      If ExisteCampo("RETAGUARDA", "ENTREGADOR", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ENTREGADOR NVARCHAR(30) "
      
      If ExisteCampo("RETAGUARDA", "MONTADOR", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD MONTADOR NVARCHAR(30) "

      If ExisteTabela("RETAGUARDA", "pk_ENTREGA", "") = False Then
         SQL = "ALTER TABLE ENTREGA ADD CONSTRAINT pk_ENTREGA PRIMARY KEY (ENTREGA_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_ENTREGA_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_ENTREGA_PEDIDO", "") = False Then
         SQL = "ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
         SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_PEDIDO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   
      If ExisteTabela("RETAGUARDA", "FK_ENTREGA_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
End Sub

Public Function TRAZ_DESCRITOR(TIPO_A As String, Codigo_N As Long) As String
   TRAZ_DESCRITOR = ""
   If Trim(TIPO_A) <> "" And Codigo_N >= 0 Then
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select descricao from DESCR "
      SQL = SQL & " where codigo >= 0 "

      If IsNumeric(Codigo_N) Then _
         SQL = SQL & " and codigo = " & Codigo_N

      If Trim(TIPO_A) <> "" Then _
         SQL = SQL & " and tipo = '" & Trim(TIPO_A) & "'"

      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         TRAZ_DESCRITOR = "" & Trim(TabDESCR.Fields("descricao").Value)

      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If
End Function

'3280-0075 SAIRO

Public Function TRAZ_TIPO_USUARIO(USUARIO_ID_N As Long) As Long
   TRAZ_TIPO_USUARIO = 0

   If Not IsNull(USUARIO_ID_N) Then
      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      SQL = "select tipo from USUARIO"
      SQL = SQL & " where usuario_id = " & USUARIO_ID_N

      'If IsNumeric(Codigo_N) Then _
         SQL = SQL & " and codigo = " & Codigo_N

      'If Trim(TIPO_A) <> "" Then _
         SQL = SQL & " and tipo_a = '" & Trim(TIPO_A) & "'"

      TabDESCR.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         If Not IsNull(TabDESCR.Fields(0).Value) Then _
            TRAZ_TIPO_USUARIO = TabDESCR.Fields(0).Value

      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If
End Function

Sub GRAVA_PRIMEIRA_DATA(DATA_A As String)
   Dia_A = Day(DATA_A)
   If Len(Dia_A) = 1 Then _
      Dia_A = "0" & Dia_A

   Mes_A = Month(DATA_A)
   If Len(Mes_A) = 1 Then _
      Mes_A = "0" & Mes_A

   Ano_A = Year(DATA_A)

   SqL2 = Dia_A & Ano_A & Mes_A

   frmCRIPTO.txtOrigem.Text = SqL2
   Call frmCRIPTO.CODIFICA

   SqL2 = frmCRIPTO.txtCripto.Text

   SQL = "update EMPRESA set "
   SQL = SQL & " par = '" & SqL2 & "'"

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

   End
End Sub

Public Sub GRAVA_PESSOA_ENDERECO()
'On Error GoTo ERRO_TRATA
'fazer normaliza�o banco
   If PESSOA_ID_N >= 0 Then _
      Exit Sub
   If ENDERECO_ID_N >= 0 Then _
      Exit Sub
   If Trim(CEP_ID_A) = "" Then _
      Exit Sub

   Dim TabPE   As New ADODB.Recordset

   If TabPE.State = 1 Then _
      TabPE.Close

   PESSOAENDERECO_ID_N = 0

   SQL = "select * from PESSOAENDERECO "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   TabPE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPE.EOF Then
      SQL = "update PESSOAENDERECO set "
      SQL = SQL & " pessoa_id = " & PESSOA_ID_N
      SQL = SQL & ", endereco_id = " & ENDERECO_ID_N
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      CONECTA_RETAGUARDA.Execute SQL
      Else
      PESSOAENDERECO_ID_N = TabPE!PESSOAENDERECO_id
      ENDERECO_ID_N = TabPE!Endereco_id

         PESSOAENDERECO_ID_N = MAX_ID("PESSOAENDERECO_ID ", "PESSOAENDERECO", "", "", "", "")

         SQL = "INSERT INTO PESSOAENDERECO (PESSOAENDERECO_ID,PESSOA_ID,ENDERECO_ID,CEP_ID) "
         SQL = SQL & " VALUES (" & PESSOAENDERECO_ID_N & "," & PESSOA_ID_N & "," & ENDERECO_ID_N & ",'" & CEP_ID_A & "')"
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabPE.State = 1 Then _
      TabPE.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GRAVA_PESSOA_ENDERECO"
End Sub

Public Sub ATUALIZA_TABELA_COMISSAO()
'On Error GoTo ERRO_TRATA

   If ExisteTabela("RETAGUARDA", "COMISSAOITEM", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE COMISSAOITEM"
   If ExisteTabela("RETAGUARDA", "COMISSAO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE COMISSAO"
   
   If ExisteTabela("RETAGUARDA", "COMISSAO", "U") = False Then
      SQL = "CREATE TABLE [dbo].[COMISSAO]("
      SQL = SQL & " [COMISSAO_ID] [BIGint] NOT NULL,"
      SQL = SQL & " [EMPRESA_ID] [int] NOT NULL,"
      SQL = SQL & " [DTINI] [datetime] NOT NULL,"
      SQL = SQL & " [DTFIM] [datetime] NOT NULL,"
      SQL = SQL & " [SITUACAO] [varchar](30) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_COMISSAO] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ("
      SQL = SQL & " [COMISSAO_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_COMISSAO_EMPRESA", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAO]  WITH CHECK ADD  CONSTRAINT [FK_COMISSAO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
      SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[COMISSAO] CHECK CONSTRAINT [FK_COMISSAO_EMPRESA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "COMISSAOITEM", "U") = False Then
      SQL = "CREATE TABLE [dbo].[COMISSAOITEM]("
      SQL = SQL & " [COMISSAO_ID] [BIGint] NOT NULL,"
      SQL = SQL & " [PEDIDO_ID] [BIGint] NOT NULL,"
      SQL = SQL & " [VENDEDOR_ID] [BIGint] NOT NULL,"
      SQL = SQL & " [PRODUTO_ID] [BIGINT] NOT NULL,"
      SQL = SQL & " [DESC_PROD] [nvarchar](80) NULL,"
      SQL = SQL & " [CLIENTE_ID] [BIGINT] NOT NULL,"
      SQL = SQL & " [CNPJCPF] [nvarchar](14) NULL,"
      SQL = SQL & " [NOME_CLI] [nvarchar](30) NULL,"
      SQL = SQL & " [NUMR_NFE] [BIGint] NULL,"
      SQL = SQL & " [NUMR_CUPOM] [BIGint] NULL,"
      SQL = SQL & " [PR_ITEM_VENDA] [float] NULL,"
      SQL = SQL & " [PR_ITEM_VAREJO] [float] NULL,"
      SQL = SQL & " [PR_ITEM_ATACADO] [float] NULL,"
      SQL = SQL & " [VALR_COMIS_PROD] [float] NULL,"
      SQL = SQL & " [VALR_COMIS_TOT] [float] NULL,"
      SQL = SQL & " [QTDE_VENDIDA] [float] NULL,"
      SQL = SQL & " [PERC_COMIS] [float] NULL,"
      SQL = SQL & " [VALOR_FATURADO] [float] NULL,"
      SQL = SQL & " CONSTRAINT [PK_COMISSAOITEM] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ("
      SQL = SQL & " [COMISSAO_ID] ASC,"
      SQL = SQL & " [PEDIDO_ID] ASC,"
      SQL = SQL & " [VENDEDOR_ID] ASC,"
      SQL = SQL & " [PRODUTO_ID] Asc "
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_COMISSAOITEM_COMISSAO", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD CONSTRAINT [FK_COMISSAOITEM_COMISSAO] FOREIGN KEY([COMISSAO_ID])"
      SQL = SQL & " References [dbo].[COMISSAO]([COMISSAO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_COMISSAO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_COMISSAOITEM_PEDIDO", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD CONSTRAINT [FK_COMISSAOITEM_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
      SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_PEDIDO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_COMISSAOITEM_VENDEDOR", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_COMISSAOITEM_VENDEDOR] FOREIGN KEY([VENDEDOR_ID])"
      SQL = SQL & " References [dbo].[VENDEDOR]([VENDEDOR_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_VENDEDOR]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_COMISSAOITEM_PRODUTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_COMISSAOITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If ExisteTabela("RETAGUARDA", "FK_COMISSAOITEM_CLIENTE", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_COMISSAOITEM_CLIENTE] FOREIGN KEY([CLIENTE_ID])"
      SQL = SQL & " References [dbo].[CLIENTE]([CLIENTE_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_CLIENTE]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "ATUALIZA_TABELA_COMISSAO"
End Sub

Public Sub ATUALIZA_TABELA_FORMAPAGTO()
   If ExisteTabela("RETAGUARDA", "FORMAPAGTO", "U") = True Then

      If ExisteCampo("RETAGUARDA", "FORMA_ID", "FORMAPAGTO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.FORMAPAGTO.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"

      If ExisteCampo("RETAGUARDA", "FORMA_ID", "ITEMLANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.ITEMLANCAMENTO.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"

      If ExisteCampo("RETAGUARDA", "CONTABILIZA", "FORMAPAGTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORMAPAGTO ADD CONTABILIZA BIT NULL"

      If ExisteCampo("RETAGUARDA", "BAIXAAUTO", "FORMAPAGTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORMAPAGTO ADD BAIXAAUTO BIT NULL"

         SQL = "update FORMAPAGTO set"
         SQL = SQL & " baixaauto = 0 "
         SQL = SQL & " where baixaauto is null"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "pk_FORMAPAGTO", "") = False Then
         SQL = "ALTER TABLE FORMAPAGTO ADD CONSTRAINT pk_FORMAPAGTO PRIMARY KEY (FORMAPAGTO_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_FORMAPAGTO_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[FORMAPAGTO] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_FORMAPAGTO_EMPRESA] "
         SQL = SQL & " FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[FORMAPAGTO] CHECK CONSTRAINT [FK_FORMAPAGTO_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteCampo("RETAGUARDA", "FORMA_ID", "caixatesorariaitem") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.caixatesorariaitem.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"
      If ExisteCampo("RETAGUARDA", "FORMA_ID", "CAIXADIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.CAIXADIAITEM.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"
   End If

   If ExisteTabela("RETAGUARDA", "TIPOVENDA", "U") = True Then
      If ExisteCampo("RETAGUARDA", "FORMA_ID", "TIPOVENDA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.TIPOVENDA.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"

      If ExisteCampo("RETAGUARDA", "CONTABILIZA", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD CONTABILIZA BIT NULL"

      If ExisteCampo("RETAGUARDA", "PERC_CARTAO_DEBITO", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD PERC_CARTAO_DEBITO FLOAT NULL"

      If ExisteCampo("RETAGUARDA", "PERC_CARTAO_CREDITO", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD PERC_CARTAO_CREDITO FLOAT NULL"

      SQL = "update TIPOVENDA set "
      SQL = SQL & " contabiliza = 1"
      SQL = SQL & " where contabiliza is null "
      CONECTA_RETAGUARDA.Execute SQL

      If ExisteTabela("RETAGUARDA", "pk_TIPOVENDA", "") = False Then
         SQL = "ALTER TABLE TIPOVENDA ADD CONSTRAINT pk_TIPOVENDA PRIMARY KEY (TIPOVENDA_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_TIPOVENDA_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[TIPOVENDA] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_TIPOVENDA_EMPRESA] "
         SQL = SQL & " FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TIPOVENDA] CHECK CONSTRAINT [FK_TIPOVENDA_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If ExisteTabela("RETAGUARDA", "FK_TIPOVENDA_FORMAPAGTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[TIPOVENDA] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_TIPOVENDA_FORMAPAGTO] "
         SQL = SQL & " FOREIGN KEY([FORMAPAGTO_ID])"
         SQL = SQL & " References [dbo].[FORMAPAGTO]([FORMAPAGTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TIPOVENDA] CHECK CONSTRAINT [FK_TIPOVENDA_FORMAPAGTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      Else
         SQL = "CREATE TABLE [dbo].[TIPOVENDA]("
         SQL = SQL & " [TIPOVENDA_ID] [int] NOT NULL,"
         SQL = SQL & " [EMPRESA_ID] [int] NOT NULL,"
         SQL = SQL & " [FORMAPAGTO_ID] [int] NOT NULL,"
         SQL = SQL & " [DESCRICAO] [nvarchar](50) NOT NULL,"
         SQL = SQL & " [PARCELA] [int] NULL,"
         SQL = SQL & " [PRAZO] [int] NULL,"
         SQL = SQL & " [PERC_JUROS] [int] NULL,"
         SQL = SQL & " [CONTABILIZA] [bit] NULL,"
         SQL = SQL & " CONSTRAINT [pk_TIPOVENDA] PRIMARY KEY CLUSTERED("
         SQL = SQL & " [TIPOVENDA_ID] Asc "
         SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[TIPOVENDA]  WITH CHECK ADD  CONSTRAINT [FK_TIPOVENDA_EMPRESA] FOREIGN KEY([EMPRESA_ID]) References [dbo].[Empresa]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[TIPOVENDA] CHECK CONSTRAINT [FK_TIPOVENDA_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[TIPOVENDA]  WITH CHECK ADD  CONSTRAINT [FK_TIPOVENDA_FORMAPAGTO] FOREIGN KEY([FORMAPAGTO_ID]) References [dbo].[FormaPagto]([FORMAPAGTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   End If
End Sub

Public Sub Centraliza_MDIChild(Formulario As Form)
  Formulario.Top = (Screen.Height) / 3 - Formulario.Height / 3
  Formulario.Left = (Screen.Width) / 2 - Formulario.Width / 2
End Sub

Public Sub OrdenaListView(ByVal lvw As MSComctlLib.ListView, ByVal Coluna_Cabecalho As MSComctlLib.ColumnHeader)
   lvw.SortKey = Coluna_Cabecalho.Index - 1
   lvw.Sorted = True
   If lvw.SortOrder = lvwAscending Then
      lvw.SortOrder = lvwDescending
      Else: lvw.SortOrder = lvwAscending
   End If
End Sub

Public Sub CRIA_IMPREL()
   If Not ExisteTabela("RETAGUARDA", "IMPREL", "U") = True Then
      SQL = "CREATE TABLE [dbo].[IMPREL]("
      SQL = SQL & " [IMPREL_ID] [bigint] NOT NULL,"
      SQL = SQL & " [RELATORIO] [nvarchar](60) NOT NULL,"
      SQL = SQL & " [CAMINHO] [nvarchar](max) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_IMPREL] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([IMPREL_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   '==================
   If Not ExisteTabela("RETAGUARDA", "PEDIDOITEMOBS", "U") = True Then
      SQL = "CREATE TABLE [dbo].[PEDIDOITEMOBS]("
      SQL = SQL & " [PEDIDO_ID] [bigint] not NULL,"
      SQL = SQL & " [SEQ_ID] [bigint] not NULL,"
      SQL = SQL & " [PRODUTO_ID] [bigint] not NULL,"

SQL = SQL & " OBS        nvarchar(MAX), "
SQL = SQL & " REFERENCIA nvarchar(MAX), "

      SQL = SQL & " [QTDE] [float] ,"
      SQL = SQL & " [Valor] [Float]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOITEMOBS]  WITH CHECK ADD  CONSTRAINT [FK_PEDIDOITEMOBS_PEDIDOITEM] FOREIGN KEY([PEDIDO_ID], [SEQ_ID])"
      SQL = SQL & " References [dbo].[PEDIDOITEM]([PEDIDO_ID], [SEQ_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PEDIDOITEMOBS] CHECK CONSTRAINT [FK_PEDIDOITEMOBS_PEDIDOITEM]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
End Sub

Public Sub CHAMA_PRODUTO_SIMPLIFICADO()
   Indr_Consulta = True
   frmCADASTROPRODUTO.txtEstoqueMaximo.Visible = False
   frmCADASTROPRODUTO.txtEstoqueMinimo.Visible = False
   frmCADASTROPRODUTO.txtPerc.Visible = False
   frmCADASTROPRODUTO.txtUN.Visible = False
   frmCADASTROPRODUTO.cmbST.Visible = False
   frmCADASTROPRODUTO.cmbALIQUOTA.Visible = False
   frmCADASTROPRODUTO.cmbOrigemMercadoria.Visible = False
   frmCADASTROPRODUTO.txtCodgNCM.Visible = False
   frmCADASTROPRODUTO.txtPercIVA.Visible = False
   frmCADASTROPRODUTO.Label23.Visible = False
   frmCADASTROPRODUTO.txtPrecoCusto.Visible = False
   frmCADASTROPRODUTO.Label25.Visible = False
   frmCADASTROPRODUTO.txtCustoAnterior.Visible = False
   frmCADASTROPRODUTO.Label26.Visible = False
   frmCADASTROPRODUTO.txtEmbalagem.Visible = False
   frmCADASTROPRODUTO.Label3.Visible = False
   frmCADASTROPRODUTO.Label5.Visible = False
   frmCADASTROPRODUTO.Label9.Visible = False
   frmCADASTROPRODUTO.Label14.Visible = False
   frmCADASTROPRODUTO.Label11.Visible = False
   frmCADASTROPRODUTO.Toolbar1.Visible = False
   frmCADASTROPRODUTO.Label12.Visible = False
   frmCADASTROPRODUTO.txtFORNEC.Visible = False
   frmCADASTROPRODUTO.Show 1
   Indr_Consulta = False
End Sub

Private Sub HtmlExportacao()
On Error GoTo CheckHtml

Dim f As ADODB.Field
Dim i As Integer

If cnn.State = 1 Then cnn.Close

cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtArquivo.Text
Rs.CursorLocation = adUseClient
Rs.Open "select * from " & cboTabela.Text, cnn, 2, 3
pb.Max = Rs.RecordCount
pb.Value = 0
Screen.MousePointer = vbHourglass

Open txtDestino.Text For Output As #1
Print #1, "<HTML>"
Print #1, "<BODY ALIGN=CENTER>"
Print #1, "<TABLE BORDER=1>"
Print #1, "<TR>"
For Each f In Rs.Fields
    Print #1, "<TD><B>" & f.Name & "</B></TD>"
Next

Print #1, "</TR>"

For i = 1 To Rs.RecordCount
   pb.Value = i
   Print #1, "<TR>"
   For Each f In Rs.Fields
      Print #1, "<TD>" & Rs.Fields(f.Name) & "</TD>"
   Next
   Print #1, "</TR>"
   Rs.MoveNext
Next
Screen.MousePointer = vbNormal
MsgBox "Convers�o realizada com sucesso"
pb.Value = 0
Rs.Close
Set Rs = Nothing
cnn.Close
Set cnn = Nothing
Print #1, "</TABLE>"
Print #1, "</BODY>"
Print #1, "</HTML>"
Close #1
Exit Sub
CheckHtml:
MsgBox Err.Description

End Sub
'4- Exportando para o formato .xml:

Private Sub XmlExportacao()
'On Error GoTo CheckXml

On Error Resume Next

Dim xmlDoc As New DOMDocument
Dim xmlroot As IXMLDOMElement
Dim node As IXMLDOMNode
Dim XMLRootTags As String, done As Boolean
Dim f As ADODB.Field
Dim i As Integer

If cnn.State = 1 Then cnn.Close

cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtArquivo.Text
Rs.CursorLocation = adUseClient
Rs.Open "select * from " & cboTabela.Text, cnn, 2, 3
pb.Max = Rs.RecordCount
pb.Value = 0
Screen.MousePointer = vbHourglass
XMLRootTags = "<Root></Root>"
done = xmlDoc.loadXML(XMLRootTags)

If done = True Then
    Set xmlroot = xmlDoc.documentElement
    For i = 1 To Rs.RecordCount
        pb.Value = i
        Set node = xmlDoc.createNode(NODE_ELEMENT, "Node" & i, "")
        xmlroot.appendChild node
        Set nodeTag = xmlroot.SelectSingleNode("Node" & i)
        For Each f In Rs.Fields
           Set node = xmlDoc.createNode(NODE_ELEMENT, f.Name, "")
           node.Text = Rs.Fields(f.Name)
           nodeTag.appendChild node
        Next
       Rs.MoveNext
   Next
Else
   MsgBox "error"
Exit Sub
End If
Screen.MousePointer = vbNormal
MsgBox "Convers�o realizada com sucesso"
pb.Value = 0
xmlDoc.Save txtDestino.Text
Rs.Close
Set Rs = Nothing
cnn.Close
Set cnn = Nothing
' Exit Sub
'CheckXml:
' MsgBox Err.Description

End Sub

Public Function Mostra_Descri��o_TipoVenda(TipoVenda_ID As Long) As String
'On Error GoTo ERRO_TRATA

   Dim TabTipoVenda As New ADODB.Recordset

   If Not IsNull(TipoVenda_ID) Then
      SQL = "select descricao from TIPOVENDA "
      SQL = SQL & " where tipovenda_id = " & TipoVenda_ID
      TabTipoVenda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTipoVenda.EOF Then _
         If Trim(TabTipoVenda.Fields(0).Value) <> "" Then _
            Mostra_Descri��o_TipoVenda = Trim(TabTipoVenda.Fields(0).Value)
   End If
   If TabTipoVenda.State = 1 Then _
      TabTipoVenda.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "Mostra_Descri��o_TipoVenda"
End Function

Public Function CONVERTE_VALOR_GRAMA(VALOR_VENDA As Double, VALOR_KILO As Double, PRODUTO_ID_N As Long) As Double
'On Error GoTo ERRO_TRATA

   Dim TabKilo As New ADODB.Recordset

   CONVERTE_VALOR_GRAMA = 0

   If PRODUTO_ID_N > 0 Then
      If TabKilo.State = 1 Then _
         TabKilo.Close

      SQL = "select preco_venda from PRODUTO "
      SQL = SQL & " where produto_id = " & PRODUTO_ID_N
      TabKilo.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabKilo.EOF Then _
         If Not IsNull(TabKilo.Fields(0).Value) Then _
            VALOR_KILO = 0 & TabKilo.Fields(0).Value

      If TabKilo.State = 1 Then _
         TabKilo.Close
   End If
'no cadastro de produto o valor de venda � gravado baseado em um kilo
   If VALOR_VENDA > 0 And VALOR_KILO > 0 Then _
      CONVERTE_VALOR_GRAMA = VALOR_VENDA / VALOR_KILO

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "CONVERTE_VALOR_GRAMA"
End Function

Public Sub RODA_AT_ESTOQUE(PROD_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select produto_id,qtde from PRODUTO"
   If PROD_ID_N > 0 Then _
      SQL = SQL & " where produto_id = " & PROD_ID_N
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF

      DoEvents

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from ESTOQUE "
      SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "insert into ESTOQUE values("
            SQL = SQL & MAX_ID("estoque_id", "estoque", "", "", "", "") 'ESTOQUE_ID
            SQL = SQL & "," & ESTABELECIMENTO_ID_N                      'ESTABELECIMENTO_ID
            SQL = SQL & "," & TabProduto.Fields("produto_id").Value     'PRODUTO_ID
            SQL = SQL & "," & tpMOEDA(TabProduto.Fields("produto_id").Value)  'QTDE_ESTOQUE
         SQL = SQL & ")"
         CONECTA_RETAGUARDA.Execute SQL
frmATUALIZACAO.Command14.Caption = TabProduto.Fields("produto_id").Value
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "RODA_AT_ESTOQUE"
End Sub

Public Sub GERA_NUMR_REQ()
'On Error GoTo ERRO_TRATA

   MOSTRA_MSG "Gerando Pedido", ""

   PEDIDO_ID_N = 1

   While PEDIDO_ID_N > 0
      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select pedido_id from PEDIDO"
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select os_id from OS"
         SQL = SQL & " where os_id = " & PEDIDO_ID_N
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabTemp.EOF Then _
            GoTo VAZA
         
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      PEDIDO_ID_N = 1 + PEDIDO_ID_N
   Wend

VAZA:

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_pedido = seq_pedido + 1 "

   'SQL = SQL & " from EMPRESA "
   'SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   'SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and seq_pedido < " & PEDIDO_ID_N
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

   'If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

   'SQL = "select seq_pedido from EMPRESA "
   'SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   'SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   'SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   'TabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   'If Not TabEmpresa.EOF Then _
      If Not IsNull(TabEmpresa.Fields(0).Value) Then _
         pedido_id_n = TabEmpresa.Fields(0).Value
   'If TabEmpresa.State = 1 Then _
      TabEmpresa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_NUMR_REQ"
End Sub

Public Sub BAIXA_ESTOQUE(NUMR_PRODUTO_ID_N As Long, NUMR_PEDIDO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If NUMR_PEDIDO_ID_N <= 0 Then _
      Exit Sub

   Dim TabVACA As New ADODB.Recordset

   If TabVACA.State = 1 Then _
      TabVACA.Close

   SQL = "SELECT PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA,empresa_id,estabelecimento_id"
   SQL = SQL & " FROM PEDIDO "
   SQL = SQL & " INNER JOIN PEDIDOITEM "
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " AND PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID"

   SQL = SQL & " where PEDIDO.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N

   SQL = SQL & " and PEDIDO.pedido_id = " & NUMR_PEDIDO_ID_N

   If numr_produto_id > 0 Then _
      SQL = SQL & " where produto_id = & " & NUMR_PRODUTO_ID_N

   TabVACA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabVACA.EOF

      SQL = "UPDATE Produto SET "
      SQL = SQL & " qtde = qtde - " & tpMOEDA(TabVACA.Fields("QTD_PEDIDA").Value)
      'SQL = SQL & ", qtde_retido = qtde_retido - " & tpMOEDA(TabVACA.Fields("QTD_PEDIDA").Value)
      SQL = SQL & ", DT_ULT_VENDA =  '" & DMA(Date) & "'"

      SQL = SQL & " where PRODUTO.empresa_id = " & TabVACA.Fields("empresa_id").Value
      SQL = SQL & " and produto_id = " & TabVACA.Fields("produto_id").Value

      CONECTA_RETAGUARDA.Execute SQL

      '================estoque
      SQL = "update ESTOQUE set "
      SQL = SQL & " QTDE_ESTOQUE = QTDE_ESTOQUE - " & tpMOEDA(TabVACA.Fields("QTD_PEDIDA").Value)

      SQL = SQL & " FROM EMPRESA "
      SQL = SQL & " INNER JOIN ESTABELECIMENTO "
      SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID "
      SQL = SQL & " INNER JOIN ESTOQUE "
      SQL = SQL & " ON ESTABELECIMENTO.ESTABELECIMENTO_ID = ESTOQUE.ESTABELECIMENTO_ID"

      SQL = SQL & " where produto_id = " & TabVACA.Fields("produto_id").Value
      SQL = SQL & " and ESTABELECIMENTO.empresa_id = " & TabVACA.Fields("empresa_id").Value
      SQL = SQL & " and ESTOQUE.estabelecimento_id = " & TabVACA.Fields("estabelecimento_id").Value

      CONECTA_RETAGUARDA.Execute SQL
      '=======================
      TabVACA.MoveNext
   Wend
   If TabVACA.State = 1 Then _
      TabVACA.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "BAIXA_ESTOQUE"
End Sub

Public Function TRAZ_QTDE_ESTOQUE(ESTAB_ID As Long, PROD_ID_N As Long) As Double
'On Error GoTo ERRO_TRATA

   Dim TabEstoque As New ADODB.Recordset

   TRAZ_QTDE_ESTOQUE = 0
   If ESTAB_ID > 0 And PROD_ID_N >= 0 Then
      If TabEstoque.State = 1 Then _
         TabEstoque.Close

      SQL = "select qtde_estoque from ESTOQUE "
      SQL = SQL & " where estabelecimento_id = " & ESTAB_ID
      SQL = SQL & " and produto_id = " & PROD_ID_N

      TabEstoque.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEstoque.EOF Then _
         If Not IsNull(TabEstoque.Fields(0).Value) Then _
            TRAZ_QTDE_ESTOQUE = TabEstoque.Fields(0).Value

      If TabEstoque.State = 1 Then _
         TabEstoque.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "TRAZ_QTDE_ESTOQUE"
End Function

'============================
Public Sub IMPRIMIR_ORDEM_SERVI�O_VEICULO(NUMR_OS_N As Long)
'On Error GoTo ERRO_TRATA

   If NUMR_OS_N <= 0 Then _
      Exit Sub

   Dim NOME_CT_A        As String
   Dim ENDERECO_EMP_A   As String
   Dim CEP_EMP_A        As String
   Dim COMP_EMP_A       As String
   Dim NUMERO_EMP_A     As String
   Dim BAIRRO_EMP_A     As String
   Dim CIDADE_EMP_A     As String
   Dim UF_EMP_A         As String
   Dim FONE_EMP_A       As String
   Dim FONE_CLIENTE_A   As String
   Dim DT_FECHA_A       As String
   Dim RESPONSAVEL_A    As String

   NOME_CT_A = ""
   ENDERECO_EMP_A = ""
   CEP_EMP_A = ""
   COMP_EMP_A = ""
   NUMERO_EMP_A = ""
   BAIRRO_EMP_A = ""
   CIDADE_EMP_A = ""
   UF_EMP_A = ""
   FONE_EMP_A = ""
   FONE_CLIENTE_A = ""
   DT_FECHA_A = ""

   SQL = "delete from OSRELITEM where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from OSREL where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOS_Servico "
   SQL = SQL & " where os_id = " & NUMR_OS_N

   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'CONSULTOR TECNICO
      
      SQL = "SELECT nome FROM USUARIO "
      SQL = SQL & " where usuario_id = " & TabTemp.Fields("ct_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            If Trim(TabConsulta.Fields(0).Value) <> "" Then _
               NOME_CT_A = "" & Trim(TabConsulta.Fields(0).Value)
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'ENDERE�O EMPRESA
      SQL = "SELECT ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
      SQL = SQL & " ENDERECO.NUMERO, CEP.Cidade, CEP.UF, CEP.CODIGO_IBGE, CEP.Cep"
      SQL = SQL & " FROM ENDERECO "
      SQL = SQL & " INNER JOIN EMPRESA "
      SQL = SQL & " ON ENDERECO.PESSOA_ID = EMPRESA.PESSOA_ID "
      SQL = SQL & " LEFT OUTER JOIN CEP "
      SQL = SQL & " ON ENDERECO.CEP = CEP.Cep"
      SQL = SQL & " Where EMPRESA.empresa_ID = " & TabTemp.Fields("empresa_id").Value
      SQL = SQL & " and endereco.tipo = 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         ENDERECO_EMP_A = "" & Trim(TabConsulta.Fields("rua").Value)
         CEP_EMP_A = "" & Trim(TabConsulta.Fields("CEP").Value)
         COMP_EMP_A = "" & Trim(TabConsulta.Fields("COMPLEMENTO").Value)
         NUMERO_EMP_A = "" & Trim(TabConsulta.Fields("NUMERO").Value)
         BAIRRO_EMP_A = "" & Trim(TabConsulta.Fields("BAIRRO").Value)
         CIDADE_EMP_A = "" & Trim(TabConsulta.Fields("cidade").Value)
         UF_EMP_A = "" & Trim(TabConsulta.Fields("uf").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE EMPRESA
      SQL = "SELECT FONE.NUMERO, FONE.DDD, FONE.LOCAL, FONE.RAMAL"
      SQL = SQL & " FROM EMPRESA "
      SQL = SQL & " INNER JOIN PESSOA "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " where empresa_id = " & TabTemp.Fields("empresa_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_EMP_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_EMP_A = FONE_EMP_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_EMP_A = FONE_EMP_A & "  "

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE cliente
      SQL = "SELECT FONE.PROP, FONE.NUMERO, FONE.DDD, FONE.LOCAL, FONE.RAMAL "
      SQL = SQL & " FROM CLIENTE "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON CLIENTE.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " Where CLIENTE_ID = " & TabTemp.Fields("cliente_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_CLIENTE_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & "  "

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      DT_FECHA_A = "" & TabTemp.Fields("DT_FECHA").Value
      If Trim(DT_FECHA_A) = "" Then
         DT_FECHA_A = 0
      End If

      SQL = "insert into OSREL "
         SQL = SQL & "("
         SQL = SQL & "OS_ID,DT_OS,TIPO_OS,SITUACAO_OS,CONSULTOR_OS,"
         SQL = SQL & "KM_OS,PLACA_OS,EMPRESA_ID,DT_OS_FEHCA,NUMR_FROTA_OS,"
         SQL = SQL & "NOME_EMP,CNPJ_EMP,ENDERECO_EMP,NUMERO_EMP,COMPLEM_EMP,"
         SQL = SQL & "CEP_EMP,BAIRRO_EMP,CIDADE_EMP,UF_EMP,FONE_EMP,NOME_CLI,"
         SQL = SQL & "CNPJCPF_CLI,FONE_CLI,DESC_VEICULO,COR_VEICULO,MARCA_VEICULO,"
         SQL = SQL & "TIPO_VEICULO,ANO_VEICULO,MODELO_VEICULO,COMB_VEICULO,"
         SQL = SQL & "CHASSI_VEICULO,MOTOR_VEICULO"
         SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & NUMR_OS_N                                                               'OS_ID
         SQL = SQL & ",'" & DMA(TabTemp.Fields("dt_os").Value) & "'"                         'DT_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("H", TabTemp.Fields("tipo_os").Value) & "'"       'TIPO_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("Z", TabTemp.Fields("SITUACAO_OS").Value) & "'"   'SITUACAO_OS
         SQL = SQL & ",'" & Trim(NOME_CT_A) & "'"                                            'CONSULTOR_OS
         SQL = SQL & "," & Trim(TabTemp.Fields("km").Value)                                  'KM_OS
         SQL = SQL & ",'" & Trim(TabTemp.Fields("PLACA").Value) & "'"                        'PLACA_OS
         SQL = SQL & "," & Trim(TabTemp.Fields("EMPRESA_ID").Value)                          'EMPRESA_ID
         SQL = SQL & ",'" & DMA(DT_FECHA_A) & "'"                      'DT_OS_FEHCA
         SQL = SQL & ",0" & Trim(TabTemp.Fields("NUMR_FROTA").Value)                          'NUMR_FROTA_OS
         SQL = SQL & ",'" & Trim(TabTemp.Fields("NOME_FANT").Value) & "'"                    'NOME_EMP
         SQL = SQL & ",'" & Trim(TabTemp.Fields("CGC").Value) & "'"                          'CNPJ_EMP
         SQL = SQL & ",'" & Trim(Replace(ENDERECO_EMP_A, ",", ".")) & "'"                    'ENDERECO_EMP
         SQL = SQL & "," & Trim(NUMERO_EMP_A)                                                'NUMERO_EMP
         SQL = SQL & ",'" & Trim(Replace(COMP_EMP_A, ",", ".")) & "'"                        'COMPLEM_EMP
         SQL = SQL & ",'" & Trim(CEP_EMP_A) & "'"                                            'CEP_EMP
         SQL = SQL & ",'" & Trim(Replace(BAIRRO_EMP_A, ",", ".")) & "'"                      'BAIRRO_EMP
         SQL = SQL & ",'" & Trim(CIDADE_EMP_A) & "'"                                         'CIDADE_EMP
         SQL = SQL & ",'" & Trim(UF_EMP_A) & "'"                                             'UF_EMP
         SQL = SQL & ",'" & Trim(FONE_EMP_A) & "'"                                           'FONE_EMP
         SQL = SQL & ",'" & Trim(TabTemp.Fields("nome").Value) & "'"                         'NOME_CLI
         SQL = SQL & ",'" & Trim(TabTemp.Fields("cgccpf").Value) & "'"                       'CNPJCPF_CLI
         SQL = SQL & ",'" & Trim(FONE_CLIENTE_A) & "'"                                       'FONE_CLI
         SQL = SQL & ",'" & Trim(TabTemp.Fields("DescricaoVeiculo").Value) & "'"             'DESC_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("S", TabTemp.Fields("cor_id").Value) & "'"        'COR_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value) & "'"      'MARCA_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("A", TabTemp.Fields("tipo_eqp").Value) & "'"      'TIPO_VEICULO
         SQL = SQL & "," & Trim(TabTemp.Fields("ano").Value)                                 'ANO_VEICULO
         SQL = SQL & "," & Trim(TabTemp.Fields("modelo").Value)                              'MODELO_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("U", TabTemp.Fields("tipo_eqp").Value) & "'"      'COMB_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("chassi").Value) & "'"                       'CHASSI_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("motor").Value) & "'"                        'MOTOR_VEICULO
      SQL = SQL & ")"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

      'ITENS SERVI�O
      SQL = "SELECT * FROM OSSERVICO "
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "SELECT nome FROM USUARIO "
         SQL = SQL & " where usuario_id = " & TabConsulta.Fields("responsavel_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)

         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSSERVICO_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'S' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,"
               SQL = SQL & "DT_CAD,DESCRICAO,VALR_ITEM,VALR_DESCONTO,QTDE,"
               SQL = SQL & " RESPONSAVEL, CODG_PRODUTO "
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSSERVICO_ID").Value              'OSRELITEM_ID
               SQL = SQL & ",'S'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("responsavel_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSTAREFA_ID").Value               'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("DESCRICAO").Value) & "'"    'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_servico").Value)    'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_servico").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(1)                                            'QTDE
               SQL = SQL & ",'" & Trim(Left(RESPONSAVEL_A, 20)) & "'"                  'RESPONSAVEL
               SQL = SQL & ",''"                                                       'CODG_PRODUTO
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'ITENS PRODUTO
      SQL = "SELECT OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO "
      SQL = SQL & " FROM OSPECA "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         NOME_A = Replace(TabConsulta.Fields("DESCRICAO").Value, ",", ".")
         NOME_A = Replace(NOME_A, "'", "�")

         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "SELECT nome_VEND FROM VENDEDOR "
         SQL = SQL & " where vendedor_id = " & TabConsulta.Fields("SOLICITANTE_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)

         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSPECA_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'P' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,"
               SQL = SQL & "DT_CAD,DESCRICAO,VALR_ITEM,VALR_DESCONTO,QTDE,RESPONSAVEL,CODG_PRODUTO"
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'OSRELITEM_ID
               SQL = SQL & ",'P'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("SOLICITANTE_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(NOME_A) & "'"                                   'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_ITEM").Value)       'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_PRODUTO").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTDE").Value)             'QTDE
               SQL = SQL & ",'" & Trim(RESPONSAVEL_A) & "'"                            'RESPONSAVEL
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("CODG_PRODUTO").Value) & "'" 'CODG_PRODUTO
            SQL = SQL & ")"
Debug.Print SQL
            CONECTA_RETAGUARDA.Execute SQL
         End If

         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   FORMULA_REL = "{OSREL.empresa_id} = " & EMPRESA_ID_N
   FORMULA_REL = FORMULA_REL & " and {OSREL.OS_ID} = " & NUMR_OS_N

   ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   Nome_Relatorio = "REL_OS.rpt"
   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlOS", "IMPRIMIR_ORDEM_SERVI�O_VEICULO"
End Sub

Public Sub IMPRIMIR_ORDEM_SERVI�O(NUMR_OS_N As Long, TIPO_REL As String)
'On Error GoTo ERRO_TRATA

   If NUMR_OS_N <= 0 Then _
      Exit Sub

   Dim NOME_CT_A        As String
   Dim CGC_A            As String
   Dim RAZAO_SOCIAL_A   As String
   Dim NOME_FANT_A      As String
   Dim ENDERECO_EMP_A   As String
   Dim CEP_EMP_A        As String
   Dim COMP_EMP_A       As String
   Dim NUMERO_EMP_A     As String
   Dim BAIRRO_EMP_A     As String
   Dim CIDADE_EMP_A     As String
   Dim UF_EMP_A         As String
   Dim FONE_EMP_A       As String
   Dim FONE_CLIENTE_A   As String
   Dim DT_FECHA_A       As String
   Dim RESPONSAVEL_A    As String

   CGC_A = ""
   RAZAO_SOCIAL_A = ""
   NOME_FANT_A = ""
   NOME_CT_A = ""
   ENDERECO_EMP_A = ""
   CEP_EMP_A = ""
   COMP_EMP_A = ""
   NUMERO_EMP_A = ""
   BAIRRO_EMP_A = ""
   CIDADE_EMP_A = ""
   UF_EMP_A = ""
   FONE_EMP_A = ""
   FONE_CLIENTE_A = ""
   DT_FECHA_A = ""

   SQL = "delete from OSRELITEM where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete from OSREL where os_id = " & NUMR_OS_N
   CONECTA_RETAGUARDA.Execute SQL

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from vwOS_Servico "
   SQL = SQL & " where os_id = " & NUMR_OS_N
   TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabTemp.EOF Then
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'CONSULTOR TECNICO
      
      SQL = "SELECT nome FROM USUARIO "
      SQL = SQL & " where usuario_id = " & TabTemp.Fields("ct_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then _
         If Not IsNull(TabConsulta.Fields(0).Value) Then _
            If Trim(TabConsulta.Fields(0).Value) <> "" Then _
               NOME_CT_A = "" & Trim(TabConsulta.Fields(0).Value)
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'ENDERE�O EMPRESA
      SQL = "SELECT ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
      SQL = SQL & " ENDERECO.NUMERO, CEP.Cidade, CEP.UF, CEP.CODIGO_IBGE, CEP.Cep"
      SQL = SQL & " FROM ENDERECO "
      SQL = SQL & " INNER JOIN EMPRESA "
      SQL = SQL & " ON ENDERECO.PESSOA_ID = EMPRESA.PESSOA_ID "
      SQL = SQL & " LEFT OUTER JOIN CEP "
      SQL = SQL & " ON ENDERECO.CEP = CEP.Cep"
      SQL = SQL & " Where EMPRESA.empresa_ID = " & EMPRESA_ID_N
      SQL = SQL & " and endereco.tipo = 'C' "
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         ENDERECO_EMP_A = "" & Trim(TabConsulta.Fields("rua").Value)
         CEP_EMP_A = "" & Trim(TabConsulta.Fields("CEP").Value)
         COMP_EMP_A = "" & Trim(TabConsulta.Fields("COMPLEMENTO").Value)
         NUMERO_EMP_A = "" & Trim(TabConsulta.Fields("NUMERO").Value)
         BAIRRO_EMP_A = "" & Trim(TabConsulta.Fields("BAIRRO").Value)
         CIDADE_EMP_A = "" & Trim(TabConsulta.Fields("cidade").Value)
         UF_EMP_A = "" & Trim(TabConsulta.Fields("uf").Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE EMPRESA
      SQL = "SELECT FONE.NUMERO, FONE.DDD, FONE.LOCAL, FONE.RAMAL"
      SQL = SQL & " FROM EMPRESA "
      SQL = SQL & " INNER JOIN PESSOA "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = PESSOA.PESSOA_ID "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON EMPRESA.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_EMP_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_EMP_A = FONE_EMP_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_EMP_A = FONE_EMP_A & "  "

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'TELEFONE cliente
      SQL = "SELECT FONE.PROP, FONE.NUMERO, FONE.DDD, FONE.LOCAL, FONE.RAMAL "
      SQL = SQL & " FROM CLIENTE "
      SQL = SQL & " INNER JOIN FONE "
      SQL = SQL & " ON CLIENTE.PESSOA_ID = FONE.PESSOA_ID"
      SQL = SQL & " Where CLIENTE.PESSOA_ID = " & TabTemp.Fields("PESSOA_id").Value
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         FONE_CLIENTE_A = "" & Trim(TabConsulta.Fields("ddd").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & " " & Trim(TabConsulta.Fields("numero").Value)
         FONE_CLIENTE_A = FONE_CLIENTE_A & "  "

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'EMPRESA
      SQL = "SELECT cgc,razao_social,nome_fant from EMPRESA"
      SQL = SQL & " Where empresa_ID = " & EMPRESA_ID_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabConsulta.EOF Then
         CGC_A = "" & Trim(TabConsulta.Fields(0).Value)
         RAZAO_SOCIAL_A = "" & Trim(TabConsulta.Fields(1).Value)
         NOME_FANT_A = "" & Trim(TabConsulta.Fields(2).Value)
      End If
      If TabConsulta.State = 1 Then _
         TabConsulta.Close


      DT_FECHA_A = "" & TabTemp.Fields("DT_FECHA").Value
      If Trim(DT_FECHA_A) = "" Then
         DT_FECHA_A = ""
         Else: DT_FECHA_A = DMA(DT_FECHA_A)
      End If

      SQL = "insert into OSREL "
         SQL = SQL & "("
         SQL = SQL & "OS_ID,DT_OS,TIPO_OS,SITUACAO_OS,CONSULTOR_OS,"
         SQL = SQL & "KM_OS,PLACA_OS,estabelecimento_ID,DT_OS_FEHCA,NUMR_FROTA_OS,"
         SQL = SQL & "NOME_EMP,CNPJ_EMP,ENDERECO_EMP,NUMERO_EMP,COMPLEM_EMP,"
         SQL = SQL & "CEP_EMP,BAIRRO_EMP,CIDADE_EMP,UF_EMP,FONE_EMP,NOME_CLI,"
         SQL = SQL & "CNPJCPF_CLI,FONE_CLI,DESC_VEICULO,COR_VEICULO,MARCA_VEICULO,"
         SQL = SQL & "TIPO_VEICULO,ANO_VEICULO,MODELO_VEICULO,COMB_VEICULO,"
         SQL = SQL & "CHASSI_VEICULO,MOTOR_VEICULO"
         SQL = SQL & ")"
      SQL = SQL & " values("
         SQL = SQL & NUMR_OS_N                                                               'OS_ID
         SQL = SQL & ",'" & DMA(TabTemp.Fields("dt_os").Value) & "'"                         'DT_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("H", TabTemp.Fields("tipo_os").Value) & "'"       'TIPO_OS
         SQL = SQL & ",'" & TRAZ_DESCRITOR("Z", TabTemp.Fields("SITUACAO_OS").Value) & "'"   'SITUACAO_OS
         SQL = SQL & ",'" & Trim(NOME_CT_A) & "'"                                            'CONSULTOR_OS
         SQL = SQL & "," & Trim(TabTemp.Fields("km").Value)                                  'KM_OS
         SQL = SQL & ",'" & Trim(TabTemp.Fields("EQUIPAMENTO_ID").Value) & "'"               'PLACA_OS
         SQL = SQL & "," & ESTABELECIMENTO_ID_N                                              'estabelecimento_ID
         SQL = SQL & ",'" & DT_FECHA_A & "'"                                                 'DT_OS_FEHCA
         SQL = SQL & ",0"                                                                    'NUMR_FROTA_OS

         SQL = SQL & ",'" & Trim(NOME_FANT_A) & "'"                                          'NOME_EMP
         SQL = SQL & ",'" & Trim(CGC_A) & "'"                                                'CNPJ_EMP

         SQL = SQL & ",'" & Trim(Replace(ENDERECO_EMP_A, ",", ".")) & "'"                    'ENDERECO_EMP
         SQL = SQL & "," & Trim(NUMERO_EMP_A)                                                'NUMERO_EMP
         SQL = SQL & ",'" & Trim(Replace(COMP_EMP_A, ",", ".")) & "'"                        'COMPLEM_EMP
         SQL = SQL & ",'" & Trim(CEP_EMP_A) & "'"                                            'CEP_EMP
         SQL = SQL & ",'" & Trim(Replace(BAIRRO_EMP_A, ",", ".")) & "'"                      'BAIRRO_EMP
         SQL = SQL & ",'" & Trim(CIDADE_EMP_A) & "'"                                         'CIDADE_EMP
         SQL = SQL & ",'" & Trim(UF_EMP_A) & "'"                                             'UF_EMP
         SQL = SQL & ",'" & Trim(FONE_EMP_A) & "'"                                           'FONE_EMP
         SQL = SQL & ",'" & Trim(TabTemp.Fields("nome_cliente").Value) & "'"                 'NOME_CLI
         SQL = SQL & ",'" & Trim(TabTemp.Fields("CNPJCPF").Value) & "'"                      'CNPJCPF_CLI
         SQL = SQL & ",'" & Trim(FONE_CLIENTE_A) & "'"                                       'FONE_CLI
         SQL = SQL & ",'" & Trim(TabTemp.Fields("nome_equipamento").Value) & "'"             'DESC_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("S", TabTemp.Fields("cor_id").Value) & "'"        'COR_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("W", TabTemp.Fields("marca_id").Value) & "'"      'MARCA_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("A", TabTemp.Fields("tipo_eqp").Value) & "'"      'TIPO_VEICULO
         SQL = SQL & "," & Trim(TabTemp.Fields("ano").Value)                                 'ANO_VEICULO
         SQL = SQL & "," & Trim(TabTemp.Fields("modelo").Value)                              'MODELO_VEICULO
         SQL = SQL & ",'" & TRAZ_DESCRITOR("U", TabTemp.Fields("tipo_eqp").Value) & "'"      'COMB_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("identificacao").Value) & "'"                'CHASSI_VEICULO
         SQL = SQL & ",'" & Trim(TabTemp.Fields("EQUIPAMENTO_ID").Value) & "'"               'MOTOR_VEICULO
      SQL = SQL & ")"

      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      CONECTA_RETAGUARDA.Execute SQL

      'ITENS SERVI�O
      SQL = "SELECT * FROM OSSERVICO "
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "SELECT nome FROM USUARIO "
         SQL = SQL & " where usuario_id = " & TabConsulta.Fields("responsavel_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)

         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSSERVICO_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'S' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,"
               SQL = SQL & "DT_CAD,DESCRICAO,VALR_ITEM,VALR_DESCONTO,QTDE,"
               SQL = SQL & " RESPONSAVEL, CODG_PRODUTO "
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSSERVICO_ID").Value              'OSRELITEM_ID
               SQL = SQL & ",'S'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("responsavel_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSTAREFA_ID").Value               'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("DESCRICAO").Value) & "'"    'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_servico").Value)    'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_servico").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(1)                                            'QTDE
               SQL = SQL & ",'" & Trim(Left(RESPONSAVEL_A, 20)) & "'"                  'RESPONSAVEL
               SQL = SQL & ",''"                                                       'CODG_PRODUTO
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If
         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close

      'ITENS PRODUTO
      SQL = "SELECT OSPECA.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO "
      SQL = SQL & " FROM OSPECA "
      SQL = SQL & " INNER JOIN PRODUTO "
      SQL = SQL & " ON OSPECA.PRODUTO_ID = PRODUTO.PRODUTO_ID"
      SQL = SQL & " where os_id = " & NUMR_OS_N
      TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      While Not TabConsulta.EOF
         NOME_A = Replace(TabConsulta.Fields("DESCRICAO").Value, ",", ".")
         NOME_A = Replace(NOME_A, "'", "�")

         If TabItem.State = 1 Then _
            TabItem.Close

         'responsavel
         RESPONSAVEL_A = ""
         SQL = "SELECT nome_VEND FROM VENDEDOR "
         SQL = SQL & " where vendedor_id = " & TabConsulta.Fields("SOLICITANTE_id").Value
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabItem.EOF Then _
            If Not IsNull(TabItem.Fields(0).Value) Then _
               If Trim(TabItem.Fields(0).Value) <> "" Then _
                  RESPONSAVEL_A = "" & Trim(TabItem.Fields(0).Value)

         If TabItem.State = 1 Then _
            TabItem.Close

         SQL = "select * from OSRELITEM "
         SQL = SQL & " where os_id = " & NUMR_OS_N
         SQL = SQL & " and osrelitem_id = " & TabConsulta.Fields("OSPECA_ID").Value
         SQL = SQL & " and TIPO_ITEM = 'P' "
         TabItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabItem.EOF Then
            SQL = "insert into OSRELITEM "
               SQL = SQL & "("
               SQL = SQL & "OS_ID,OSRELITEM_ID,TIPO_ITEM,USU_ID,PROSERV_ID,"
               SQL = SQL & "DT_CAD,DESCRICAO,VALR_ITEM,VALR_DESCONTO,QTDE,RESPONSAVEL,CODG_PRODUTO"
               SQL = SQL & ")"
            SQL = SQL & " values("
               SQL = SQL & NUMR_OS_N                                                   'OS_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'OSRELITEM_ID
               SQL = SQL & ",'P'"                                                      'TIPO_ITEM
               SQL = SQL & "," & TabConsulta.Fields("SOLICITANTE_ID").Value            'USU_ID
               SQL = SQL & "," & TabConsulta.Fields("OSPECA_ID").Value                 'PROSERV_ID
               SQL = SQL & ",'" & DMA(TabConsulta.Fields("dt_cad").Value) & "'"        'DT_CAD
               SQL = SQL & ",'" & Trim(NOME_A) & "'"                                   'DESCRICAO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("valor_ITEM").Value)       'VALR_ITEM
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("desconto_PRODUTO").Value) 'VALR_DESCONTO
               SQL = SQL & "," & tpMOEDA(TabConsulta.Fields("QTDE").Value)             'QTDE
               SQL = SQL & ",'" & Trim(RESPONSAVEL_A) & "'"                            'RESPONSAVEL
               SQL = SQL & ",'" & Trim(TabConsulta.Fields("CODG_PRODUTO").Value) & "'" 'CODG_PRODUTO
            SQL = SQL & ")"

            CONECTA_RETAGUARDA.Execute SQL
         End If

         If TabItem.State = 1 Then _
            TabItem.Close

         TabConsulta.MoveNext
      Wend
      If TabConsulta.State = 1 Then _
         TabConsulta.Close
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close

   FORMULA_REL = "{OSREL.estabelecimento_id} = " & ESTABELECIMENTO_ID_N
   FORMULA_REL = FORMULA_REL & " and {OSREL.OS_ID} = " & NUMR_OS_N

   ESCOLHE_IMPRESSORA NOME_BANCO_DADOS

   If Trim(TIPO_REL) = "OFICINA" Then _
      Nome_Relatorio = "REL_OFICINA.rpt"
   If Trim(TIPO_REL) = "SERVI�O" Then _
      Nome_Relatorio = "REL_SERVICO.rpt"

   frmRELATORIO10.Show 1

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlOS", "IMPRIMIR_ORDEM_SERVI�O"
End Sub

Public Sub MATA_FONE(NUMERO_A As String, ID_PESSOA As Long, CNPJ_CPF_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(NUMERO_A) <> "" Then

      SQL = "delete FONE "
      SQL = SQL & " where numero = '" & Trim(NUMERO_A) & "'"

      If Trim(CNPJ_CPF_A) <> "" Then _
         SQL = SQL & " and prop = '" & Trim(CNPJ_CPF_A) & "'"

      If IsNumeric(ID_PESSOA) Then _
         If ID_PESSOA > 0 Then _
            SQL = SQL & " and pessoa_id = " & ID_PESSOA

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlOS", "MATA_FONE"
End Sub
