Attribute VB_Name = "modTEF"
'Option Explicit
'Public iConta As Integer
'Public cHora As String
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
''Declaraciones para 32 bits
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'    (ByVal hwnd As Long, ByVal wMsg As Long, _
'    ByVal wParam As Long, lParam As Any) As Long
'
'Public Const CB_SHOWDROPDOWN = &H14F
'
'Dim lFlag As Boolean
'Public iTransacao As Integer
'Dim i As Long
'Dim response As Integer
'Dim linhaArquivo As String
'Public naoConfirmado As Boolean
'Dim CodigoFilial As Integer
'Dim rs As New adodb.Recordset
'
'Public Sub BuscaInformacoesEmpresa()

'
''    rs.Open "SELECT CodigoFilial, isnull(NFeletronica,'N') as NFeletronica, isnull(CaminhoGerenciadorPadrao, 'C:\TEF_DIAL\') as CaminhoGerenciadorPadrao, isnull(EmiteCupomFiscal, 'N') as EmiteCupomFiscal FROM Filial WHERE (CodigoEmpresa = " & intEmpresa & ") AND (CodigoFilial = (SELECT CodigoFilial FROM Funcionario WHERE CodigoEmpresa = " & intEmpresa & " and CodigoFuncionario = " & intUsuario & "))", CONECTA_RETAGUARDA , , , adCmdText
''    If Not rs.EOF Then
''        intCodigoFilial = rs!CodigoFilial
''        strEmiteCupomFiscal = rs!EmiteCupomFiscal
''        strCaminhoGerenciadorPadrao = rs!CaminhoGerenciadorPadrao
''        If rs!NFeletronica = "S" Then
''            booNFeletronica = rs!NFeletronica
''        Else
''            booNFeletronica = rs!NFeletronica
''        End If
''    Else
''        booNFeletronica = "N"
''    End If
''    rs.Close
''
''    PAR_LOCAL_VendaCartao = True
''
''    mdiPrincipal.MousePointer = 0
'
'    Exit Sub
'ERRO_TRATA:
'    'mdiPrincipal.MousePointer = 0
'    'ControleErros Err.Number, Err.Description, Err.Source, "ImportaNotaFiscal"
'End Sub
'
Public Function AprovaTEF(intFilial As Integer, NumeroCupon As String, Valor As Currency) As Boolean
On Error GoTo ERRO_TRATA
    Dim iArquivo As Integer
    Dim strLeitura As String

    Dim cValorPago As String
    'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

    CodigoFilial = intFilial
    iConta = 1

    'verifica se existe o diretorio
    If Dir(App.Path & "\" & Format(CodigoFilial, "00")) = "" Then
        MkDir App.Path & "\" & Format(CodigoFilial, "00") 'caso não exista cria diretório
    End If

    'Se já existe mais de um pagamento, deve-se confirmar a transação anterior
    If Dir(App.Path & "\" & Format(CodigoFilial, "00") & "\PENDENTE.TXT") <> "" Then
        iArquivo = FreeFile
        Open App.Path & "\" & Format(CodigoFilial, "00") & "\PENDENTE.TXT" For Input As iArquivo
        Line Input #iArquivo, strLeitura
        If IsNumeric(CInt(Trim(strLeitura))) Then
            Close iArquivo
            ConfirmaTransacao (CInt(Trim(strLeitura)))
            MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\PENDENTE.TXT")
        End If
    End If

    'Pega a hora atual
    cHora = Time
    cValorPago = tpMOEDA(Valor & "")
    cValorPago = Replace(cValorPago, ",", "", , , vbTextCompare)

    intRetorno = RealizaTransacao(CDate(cHora), NumeroCupon, cValorPago, iConta)
    If intRetorno = 1 Then
        intRetorno = Bematech_FI_EfetuaFormaPagamento("CARTAO", Format$(Valor, "Standard"))
        'Call VerificaRetornoImpressoraBematech("", "", "Efetua pagamento")
        AprovaTEF = True
    Else
        AprovaTEF = False
    End If

    Exit Function
ERRO_TRATA:
    If Err.Number = 75 Then 'ja existe o diretório
        Resume Next
    Else
        'ControleErros Err.Number, Err.Description, Err.Source, ""
   End If
End Function
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    VerificaGerenciadorPadrao
''// Objetivo:
''//    Verificar se o Gerenciador Padrão está ativo
''// Parâmetro:
''//    não há
''// Retorno:
''//    True para Gerenciador Padrão ATIVO
''//    False para Gerenciador Padrão INATIVO
''//
''////////////////////////////////////////////////////////////////////////////////
Function VerificaGerenciadorPadrao(CodigoFilial As Integer) As Boolean
On Error GoTo ERRO_TRATA

    Dim cConteudoArquivo As String
    Dim hora As Date
    Dim iTentativas As Integer

    'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

    lFlag = True

    'verifica se existe o diretorio
    If Dir(App.Path & "\" & Format(CodigoFilial, "00")) = "" Then
        MkDir App.Path & "\" & Format(CodigoFilial, "00") 'caso não exista cria diretório
    End If

IniciarDeNovo:
    hora = Date & " " & Time
    cConteudoArquivo = ""
    cConteudoArquivo = "000-000 = ATV" & vbCrLf & _
              "001-000 = " & hora & _
              "999-999 = 0"
    Call GravaArquivo_Binario(App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", cConteudoArquivo)

   ' Copia o arquivo para o diretório do Gerenciador Padrão
    FileCopy App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001"

    ' Apaga o arquivo local
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001")

    iTentativas = 1
    For iTentativas = 1 To 7 Step 1
        If Dir(strCaminhoGerenciadorPadrao & "\RESP\ATIVO.001") <> "" Or Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS") <> "" Then
            'O Gerencial Padrão se encontra ATIVO
            iTentativas = 8
            lFlag = True
            Sleep (1000)
            VerificaGerenciadorPadrao = True
            Exit For
        End If

        Sleep (1000)
        If iTentativas = 7 Then
            'O Gerencial Padrão se encontra INATIVO
            iTentativas = 8
            lFlag = False
            VerificaGerenciadorPadrao = False
            Exit For
        End If
    Next iTentativas

    If lFlag = False Then
        Shell strCaminhoGerenciadorPadrao & "\tef_dial.exe", vbNormalFocus
        'SuperShell strCaminhoGerenciadorPadrao & "\tef_dial.exe", strCaminhoGerenciadorPadrao, 0, SW_NORMAL, HIGH_PRIORITY_CLASS

        Call TEFMensagemPopup("Gerenciador Padrão não está ativo", "e será ativado automaticamente.", "Caminho GP:" & strCaminhoGerenciadorPadrao, "Tecban: " & IIf(iTEFTecban <> 0, "SIM", "NÃO"))

        GoTo IniciarDeNovo
    End If
    Exit Function
ERRO_TRATA:
    If Err.Number = 75 Then
        Resume Next
    End If
End Function
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    RealizaTransacao
''// Objetivo:
''//    Realiza a transação TEF
''// Parâmetros:
''//   TDateTime para identificar o número da transação
''//   String para o Número do Cupom Fiscal (COO)
''//   String para a Valor da Forma de Pagamento
''//   Integer com o número da transação
''// Retorno:
''//    True para OK
''//    False para não OK
''//
''////////////////////////////////////////////////////////////////////////////////
Function RealizaTransacao(hora As Date, cNumeroCupom As String, _
                           cValorPago As String, iConta As Integer) As Integer
On Error GoTo ERRO_TRATA

    Dim cConteudoArquivo As String
    Dim cLinhaArquivo As String
    Dim cLinha As String
    Dim cCampoArquivo As String
    Dim iArquivo As Integer
    Dim arquivoIncorreto As Boolean
    Dim lFlag As Boolean
    Dim iTentativas As Integer
    Dim iVezes As Integer

    Dim bTransacao As Boolean
    Dim bFlagArq As Integer
    Dim lNumeroLinha As Long
    Dim iAux As Integer
    Dim intContador As Integer

    'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

    arquivoIncorreto = True
    intContador = 1
VOLTAR:
    '''''''''''''''CRIANDO A SOLICITAÇÃO DA TRANSAÇÃO TEF'''''''''''''''''
    ' Conteúdo do arquivo INTPOS.001 para solicitar a transação TEF.
    cConteudoArquivo = ""
    cConteudoArquivo = "000-000 = CRT" & vbCrLf & _
                       "001-000 = " & Format(hora, "HhNnSs") & vbCrLf & _
                       "002-000 = " & cNumeroCupom & vbCrLf & _
                       "003-000 = " & cValorPago & vbCrLf & _
                       "999-999 = 0"
    Call GravaArquivo_Binario(App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", cConteudoArquivo)
    ' Copia o arquivo para o diretório do Gerenciador Padrão
    FileCopy App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001"
    ' Apaga o arquivo local
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Se já existe um IMPRIME[conta].TXT, deleta ele
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT")

    RealizaTransacao = -2
    'Enquanto o gerenciador padrão não tiver mandado resposta, fica em loop
    'Excedendo 7 segundos, sai da função retornando 0
    While Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS") = "" ' Verifica o arquivo INTPOS.001 de resposta.
        DoEvents
        Sleep (1000)
        iTentativas = iTentativas + 1
        If iTentativas > 7 Then
            Sleep (1000)
            'Mensagem "Gerenciador Padrão não está ativo e será ativado automaticamente!"
            iTentativas = 1
            intContador = intContador + 1
            If intContador <= 3 Then
                GoTo VOLTAR
            End If

            RealizaTransacao = 0
            Exit Function
        End If
    Wend

    lNumeroLinha = 0
    cLinhaArquivo = ""
    cLinha = ""
    Do
        While Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001") = ""  ' Verifica o arquivo INTPOS.001 de resposta.
            'Mensagem "AGUARDANDO RESPOSTA DO TEF"
            DoEvents
        Wend

        'verifica se o arquivo é valido
        iArquivo = FreeFile
        Open strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001" For Input As iArquivo

        While Not EOF(iArquivo)
            Line Input #iArquivo, cLinhaArquivo 'Lê uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo

            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
            If (cCampoArquivo = "001") Then
                If Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10) = Format(hora, "HhNnSs") Then
                    arquivoIncorreto = False
                End If
            End If
        Wend
        Close iArquivo
        If arquivoIncorreto Then
            'Mensagem "tem que tirar: Mata arquivo> " & strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001"
            MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001")
        End If

    Loop While arquivoIncorreto

    While (RealizaTransacao = -2) 'FOR1-IF1-WHILE1

        iArquivo = FreeFile
        Open strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001" For Input As iArquivo

        While Not EOF(iArquivo) 'FOR1-IF1-WHILE1-IF1-DOWHILE1
            Line Input #iArquivo, cLinhaArquivo 'Lê uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
            lNumeroLinha = lNumeroLinha + 1
            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)

            Select Case CInt(cCampoArquivo) 'FOR1-IF1-WHILE1-IF1-WHILE1-SELECT1
                Case 9: ' Verifica se a Transação foi Aprovada.
                    If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) = "0" Then
                        bTransacao = True
                        RealizaTransacao = 1
                    End If
                    If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> "0" Then
                        bTransacao = False
                        RealizaTransacao = -1
                    End If
                Case 28: ' Verifica se existem linhas para serem impressas.
                    If (CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> 0) And (bTransacao = True) Then
                        'É realizada uma cópia temporária do arquivo INTPOS.001 para cada transação efetuada.
                        'Caso a transação necessite ser cancelada, as informações estarão neste arquivo.
                         ' Copia o arquivo para o diretório do Gerenciador Padrão
                        'Se está aberto, fecha para copiar


                        Close iArquivo 'fecha arquivo
                        'FileCopy strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001", strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001"

                        RealizaTransacao = 1
                        iArquivo = FreeFile
                        Open strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001" For Input As iArquivo
                        While bFlagArq = False
                            Line Input #iArquivo, cLinhaArquivo
                            If Mid(cLinhaArquivo, 1, 3) = 28 Then
                                bFlagArq = True
                            End If
                        Wend
                        For iVezes = 1 To CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) Step 1
                            Line Input #iArquivo, cLinhaArquivo 'Lê uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
                            If Mid(cLinhaArquivo, 1, 3) = "029" Then
                                cLinha = cLinha + Mid(cLinhaArquivo, 12, Len(cLinhaArquivo) - 12) + vbCrLf
                            End If
                        Next iVezes
                    End If

                Case 30: ' Verifica se o campo é o 030 para mostrar a mensagem para o operador
                    If cLinha <> "" Then
                        'Mensagem  Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                    Else
                        MataArquivo (strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001")
                        'Mensagem  Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                        RealizaTransacao = -1
                    End If
                End Select 'FOR1-IF1-WHILE1-IF1-WHILE1-ENDSELECT1
        Wend

    Wend
        ' Cria o arquivo temporário IMPRIME.TXT com a imagem do comprovante
        If (cLinha <> "") Then
            Close iArquivo
            Call GravaArquivo_Binario(App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT", cLinha)
        End If

        Sleep (1000)
        ' O arquivo INTPOS.STS não retornou em 7 segundos, então o operador é informado.
        If (iTentativas = 7) Then
            If Dir(strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001") <> "" Then
                MataArquivo (strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001")
                'Mensagem "Gerenciador Padrão não está ativo!"
                RealizaTransacao = 0
                Exit Function
            End If
        End If
        If (RealizaTransacao = 0) Or (RealizaTransacao = -1) Then
            Close iArquivo
        Else
            RealizaTransacao = 1
            Call GravaArquivo_Binario(App.Path & "\" & Format(CodigoFilial, "00") & "\PENDENTE.TXT", Trim(CStr(iConta)))
        End If

    MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS")
    'MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001")

    Exit Function
ERRO_TRATA:
    If Err.Number = 70 Then
        Sleep 1000
        'Mensagem "Erro: " & Err.Description
    End If
    'ControleErros Err.Number, Err.Description, Err.Source, ""
End Function
'
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    ImprimeTransacao
''// Objetivo:
''//    Realiza a impressão da Transação TEF
''// Parâmetros:
''//   String para a Forma de Pagamento
''//   String para a Valor da Forma de Pagamento
''//   String para o Número do Cupom Fiscal (COO)
''//   TDateTime para identificar o número da transação
''//   Integer com o número da transação
''// Retorno:
''//    True para OK
''//    False para não OK
''//
''////////////////////////////////////////////////////////////////////////////////
Function ImprimeTransacao(ByVal cFormaPGTO As String, ByVal cValorPago As String, _
                          ByVal cCOO As String, ByVal hora As String, _
                          ByVal iConta As Integer, ByVal Gerencial As Boolean) As Boolean
On Error GoTo ERRO_TRATA

    Dim cLinhaArquivo As String
    Dim cLinha  As String
    Dim cSaltaLinha As String
    Dim iArquivo As Integer
    Dim iVezes As Integer
    Dim iRetorno As Integer
    Dim tipoImpressora As Integer
    Dim via As Integer

'   Neste ponto é criado o arquivo TEF.TXT, indicando que há uma operação de TEF sendo
'   realizada. Caso ocorra uma queda de energia, no momento da impressão do TEF, e a
'   aplicação for inicializada, ao identificar a existência deste arquivo, a transação do TEF
'   deverá ser concelada.
    'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

    Call GravaArquivo_Binario(App.Path & "\" & Format(CodigoFilial, "00") & "\TEF.TXT", CStr(iTransacao))
    'iRetorno = Bematech_FI_IniciaModoTEF()

    ImprimeTransacao = False
    If Trim(cCOO) = "" Then
        MsgBox "Não foi possível obter o número do comprovante."
'        Call Bematech_FI_FinalizaModoTEF
        If (ImprimeGerencial(iConta) = 1) Then
            ImprimeTransacao = True
            Exit Function
        Else
            Exit Function
        End If
    End If
    If Dir(App.Path + "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT") <> "" Then
        DoEvents

        ' Função para bloqueio do teclado e mouse
        'iRetorno = Bematech_FI_IniciaModoTEF()
        'iRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado

        If Not Gerencial Then
            iRetorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado(cFormaPGTO, cValorPago, cCOO)
            If Not VerificaRetornoFuncaoImpressora(iRetorno) Then
                Exit Function
            End If
        End If

        cLinha = ""
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '          INÍCIO DA LEITURA DE ARQUIVO PARA IMPRESSÃO          '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For via = 1 To 2 Step 1
            iArquivo = FreeFile
            Open App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT" For Input As iArquivo

            While Not EOF(iArquivo)
            '''''''''''''Lê uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
                Line Input #iArquivo, cLinhaArquivo

                'A função de impressão não aceita strings vazias
                If cLinhaArquivo = "" Then
                    cLinhaArquivo = " "
                End If

                '''''''''''''Imprime o que foi lido
                If Gerencial Then
                    iRetorno = Bematech_FI_RelatorioGerencial(cLinhaArquivo & vbCrLf)
                Else
                    iRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(cLinhaArquivo & vbCrLf)
                End If

                '''''''''''''Aqui é feito o tratamento de erro de comunicação com a impressora
                '''''''''''''(desligamento da impressora durante a impressão do comprovante).
                If Not (VerificaRetornoFuncaoImpressora(iRetorno)) Then
                    Close iArquivo
                    'iRetorno = Bematech_FI_FinalizaModoTEF()
                    'iRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
                    ImprimeTransacao = False
                    Exit Function
                End If
            Wend



            '''''''''''''Aciona o corte de papel
            If via = 1 Then
                '''''''''''''Pula 7 linhas
                cSaltaLinha = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
                iRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(cSaltaLinha)
                iRetorno = Bematech_FI_VerificaTipoImpressora(tipoImpressora)
                If ((tipoImpressora = "2") Or (tipoImpressora = "4") Or (tipoImpressora = "6") Or (tipoImpressora = "8")) Then
                    'iRetorno = Bematech_FI_AcionaGuilhotinaMFD(0)
                End If
                '''''''''''''Exibe mensagem na tela
                'Mensagem "Por favor, destaque a " & via & "ª via."
                Sleep (3000)
            End If

            Close iArquivo
            'Mensagem ""
        Next via
        Close iArquivo
        'iRetorno = Bematech_FI_FinalizaModoTEF()
        iRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado()
        'Mensagem "Por favor, destaque a " & (via - 1) & "ª via."

        Sleep (3000)
        'Mensagem ""
        ImprimeTransacao = True
    End If

    'Desbloqeia o teclado e o mouse
    'iRetorno = Bematech_FI_FinalizaModoTEF()

    MataArquivo (App.Path + "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT")

    Exit Function
ERRO_TRATA:
    If Err.Number = 70 Then
        Sleep 1000
        Resume
    End If
    'ControleErros Err.Number, Err.Description, Err.Source, ""
End Function
'
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    ConfirmaTransacao
''// Objetivo:
''//    Confirmar a Transação TEF
''// Parâmetros:
''//   Integer com o número da transação
''// Retorno:
''//    True para OK
''//    False para não OK
''//
''////////////////////////////////////////////////////////////////////////////////
Function ConfirmaTransacao(iConta As Integer) As Boolean
   Dim cLinhaArquivo As String
   Dim cConteudo As String
   Dim iArquivo As Integer
   Dim lFlag As Boolean
   Dim iVezes As Integer

   'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

   cLinhaArquivo = ""
   cConteudo = ""

    If Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001") <> "" Then
        If (iConta <> 0) Then
            iArquivo = FreeFile
            Open strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001" For Binary As iArquivo
        Else
            iArquivo = FreeFile
            Open strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001" For Binary As iArquivo
        End If
        While Not EOF(iArquivo)
            DoEvents
            On Error GoTo FimArquivo
            Line Input #iArquivo, cLinhaArquivo
            If (Mid(cLinhaArquivo, 1, 3) = "001") Or (Mid(cLinhaArquivo, 1, 3) = "002") Or (Mid(cLinhaArquivo, 1, 3) = "010") Or (Mid(cLinhaArquivo, 1, 3) = "012") Or (Mid(cLinhaArquivo, 1, 3) = "027") Then
                cConteudo = cConteudo & cLinhaArquivo & vbCrLf
            End If
            If (Mid(cLinhaArquivo, 1, 3) = "999") Then
                  cConteudo = cConteudo & cLinhaArquivo
            End If
FimArquivo: Wend
        Close iArquivo

        cConteudo = "000-000 = CNF" & vbCrLf & cConteudo
        Call GravaArquivo_Binario(App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", cConteudo)
        FileCopy App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001"
        MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001")
        While Not Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS") <> ""
            DoEvents
            Sleep (1000)
        Wend

        MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS")
    End If

    'Se o arquivo TEF.TXT, que identifica que houve uma transação impressa
    'existir, o mesmo será exluído.
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\TEF.TXT")

End Function
'
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    NaoConfirmaTransacao
''// Objetivo:
''//    Não Confirmar a Transação TEF
''// Parâmetros:
''//   Integer com o número da transação
''// Retorno:
''//    True para OK
''//    False para não OK
''//
''////////////////////////////////////////////////////////////////////////////////
Function NaoConfirmaTransacao(ByVal iConta As Integer) As Boolean
    Dim cLinhaArquivo As String
    Dim cConteudo As String
    Dim cCampoArquivo As String
    Dim iArquivo As Integer
    Dim lFlag As Boolean
    Dim cValor As String
    Dim cNomeRede As String
    Dim cNSU As String
    Dim cIdent As String
    Dim cData As String
    Dim cHora As String
    Dim iVezes As Integer

    'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME" + CStr(iConta) + ".TXT")
    cLinhaArquivo = ""
    cConteudo = ""

    'Se achou o INTPOS[conta].001 na pasta C:\TEF_DIAL\RESP
    If Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001") <> "" Then
        iArquivo = FreeFile
        Open strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001" For Input As iArquivo
        While Not EOF(iArquivo)
            DoEvents
            Line Input #iArquivo, cLinhaArquivo
            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
            Select Case CInt(cCampoArquivo)
                Case 1:
                    cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                Case 3:
                    cValor = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                Case 10:
                      cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                      cNomeRede = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                Case 12:
                    cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                    cNSU = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                Case 27:
                    cConteudo = cConteudo & cLinhaArquivo & vbCrLf
                Case 999:

                cConteudo = cConteudo & cLinhaArquivo
             End Select
        Wend
        Close iArquivo

        cConteudo = "000-000 = NCN" & vbCrLf & cConteudo
        iArquivo = FreeFile

        Open App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001" For Output As iArquivo
        Print #iArquivo, cConteudo
        Close iArquivo

        FileCopy App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001"
        MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001")

        While Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS") = ""
            DoEvents
            Sleep (1000)
        Wend

        MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS")

        'Se o arquivo TEF.TXT, que identifica que houve uma transação impressa
        'existir, o mesmo será exluído.
        MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\TEF.TXT")
'        mdiPrincipal.MousePointer = vbDefault
        MsgBox "Cancelada a Transação" & vbCrLf & vbCrLf & "Rede: " & _
            cNomeRede & vbCrLf & "Doc Nº: " & cNSU & vbCrLf & "Valor: " & _
            Format(CDbl(cValor) / 100, "#,##0.00"), vbOKOnly + vbInformation, _
            "Atenção"
'        mdiPrincipal.MousePointer = vbHourglass
        MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001")
        Call Bematech_FI_FechaRelatorioGerencial
        iConta = iConta - 1
        If iConta > 0 Then
            For iVezes = 1 To iConta Step 1
                If Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS" + CStr(iVezes) + ".001") <> "" Then
                    cLinhaArquivo = ""
                    cConteudo = ""
                    iArquivo = FreeFile
                    Open strCaminhoGerenciadorPadrao & "\RESP\INTPOS" & CStr(iVezes) & ".001" For Input As iArquivo
                        While Not EOF(iArquivo)
                            DoEvents
                            Line Input #iArquivo, cLinhaArquivo
                            cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
                            Select Case CInt(cCampoArquivo)
                                Case 1:
                                    cIdent = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 3:
                                    cValor = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 10:
                                    cNomeRede = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 12:
                                    cNSU = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 22:
                                    cData = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                Case 23:
                                    cHora = Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                            End Select
                        Wend
                        Close iArquivo
                        MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS")
                        Call CancelaTransacaoTEF(cNSU, cValor, cNomeRede, cNSU, cData, cHora, iVezes)
                        ConfirmaTransacao (iVezes)
                        Call Bematech_FI_FechaRelatorioGerencial
                        ImprimeGerencial (iVezes)
                        MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS")
                        ' Se o arquivo TEF.TXT, que identifica que houve uma transação impressa
                        ' existir, o mesmo será excluído.
                        MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\TEF.TXT")
                End If
            Next iVezes
        End If

        If iConta > 0 Then
            For iVezes = 1 To iConta Step 1
                MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS" & CStr(iVezes) & ".001")
                MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\CANCEL" & CStr(iVezes) & ".001")
                MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT")
                naoConfirmado = True
           Next iVezes
        End If
    End If
End Function
'
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    CancelaTransacaoTEF
''// Objetivo:
''//    Cancelar uma transação já confirmada
''// Parâmetros:
''//    String com o número de identificação (NSU)
''//    String com o valor da transação
''//    String com o valor da transação
''//    String com o nome e bandeira (REDE)
''//    String com o número do documento
''//    String com a data da transação no formato DDMMAAAA
''//    String com a hora da transação no formato HHSMMSS
''//    Integer com o número da transação
''// Retorno:
''//    True para OK
''//    False para não OK
''//
''////////////////////////////////////////////////////////////////////////////////
Function CancelaTransacaoTEF(ByVal cNSU As String, ByVal cValor As String, ByVal cNomeRede As String, _
         ByVal cNumeroDOC As String, ByVal cData As String, ByVal cHora As String, ByVal iVezes As Integer) As Boolean
    Dim cConteudo As String
    Dim iArquivo As Integer
    Dim lFlag As Boolean

    'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

    cConteudo = ""
    cConteudo = "000-000 = CNC" & vbCrLf & _
                "001-000 = " & cNSU & vbCrLf & _
                "003-000 = " & cValor & vbCrLf & _
                "010-000 = " & cNomeRede & vbCrLf & _
                "012-000 = " & cNumeroDOC & vbCrLf & _
                "022-000 = " & cData & vbCrLf & _
                "023-000 = " & cHora & vbCrLf & _
                "999-999 = 0"
    iArquivo = FreeFile
    Open App.Path + "\INTPOS.001" For Output As iArquivo

    Print #iArquivo, cConteudo
    Close iArquivo
    FileCopy App.Path + "\INTPOS.001", strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001"
    MataArquivo (App.Path + "\INTPOS.001")

    While Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001") = ""
        Sleep (1000)
    Wend

    MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS")
    FileCopy strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001", strCaminhoGerenciadorPadrao & "\RESP\CANCEL" & CStr(iVezes) & ".001"
    MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001")

End Function
'
''////////////////////////////////////////////////////////////////////////////////
''// Função:
''//    FuncaoAdministrativaTEF
''// Objetivo:
''//    Chamar o módulo administrativo da bandeira
''// Parâmetro:
''//    String com o identificador
''// Retorno:
''//    1 para OK
''//    diferente de 1 para não OK
''////////////////////////////////////////////////////////////////////////////////
'Function FuncaoAdministrativaTEF(ByVal hora As String) As Integer
'    If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa
'
'    Dim iArquivo As Integer
'    Dim lFlag As Boolean
'    Dim cConteudoArquivo As String
'
'    'Conteúdo do arquivo INTPOS.001 para solicitar a transação TEF
'    cConteudoArquivo = ""
'    cConteudoArquivo = "000-000 = ADM" & vbCrLf & _
'                       "001-000 = " & Format(hora, "HhNnSs") & vbCrLf & _
'                       "999-999 = 0"
'    Call GravaArquivo_Binario(App.Path + "\INTPOS.001", cConteudoArquivo)
'
'    FileCopy App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001", strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001"
'    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\INTPOS.001")
'
'End Function
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    ImprimeGerencial
''// Objetivo:
''//    Imprimir através do Relatório Gerencial a transação efetuada.
''// Parâmetro:
''//    Integer com o número da transação
''// Retorno:
''//    1 para OK
''//    diferente de 1 para não OK
''//
''////////////////////////////////////////////////////////////////////////////////
Function ImprimeGerencial(ByVal iConta As Integer) As Integer
    Dim iArquivo As Integer
    Dim iTentativas As Integer
    Dim iVezes As Integer
    Dim iRetorno As Integer
    Dim via As Integer
    Dim tipoImpressora As Integer
    Dim bTransacao As Boolean
    Dim cArquivoTexto As String
    Dim cArquivoIntPos As String
    Dim cArquivoCancel As String
    Dim cCampoArquivo As String
    Dim cLinha As String
    Dim cSaltaLinha As String
    Dim cLinhaArquivo As String

    'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

    If iConta = 0 Then
        cArquivoTexto = "IMPRIME.TXT"
        cArquivoIntPos = "INTPOS.001"
    Else
        cArquivoTexto = "IMPRIME.TXT"
        cArquivoIntPos = "INTPOS.001"
        cArquivoCancel = "CANCEL.001"
    End If
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\" & cArquivoTexto)

    If Dir(strCaminhoGerenciadorPadrao & "\RESP\" & cArquivoCancel) <> "" Then
        cArquivoIntPos = "CANCEL.001"
    End If
    ImprimeGerencial = -2

    For iTentativas = 1 To 7 Step 1
        cLinhaArquivo = ""
        cLinha = ""
        While (ImprimeGerencial = -2)
            If Dir(strCaminhoGerenciadorPadrao & "\RESP\" & cArquivoIntPos) <> "" Then
                iArquivo = FreeFile
                Open strCaminhoGerenciadorPadrao & "\RESP\" & cArquivoIntPos For Input As iArquivo
                While Not EOF(iArquivo)
                    Line Input #iArquivo, cLinhaArquivo
                    cCampoArquivo = Mid(cLinhaArquivo, 1, 3)
                    Select Case CInt(cCampoArquivo)
                        Case 9: ' Verifica se a Transação foi Aprovada
                            If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) = "0" Then
                                bTransacao = True
                                ImprimeGerencial = 1
                            End If
                            If (Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> "0" Then
                                bTransacao = False
                                ImprimeGerencial = -1
                            End If

                        Case 28: 'Verifica se existem linhas para serem impressas
                            If (CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) <> 0) And (bTransacao = True) Then
                                ImprimeGerencial = 1
                                For iVezes = 1 To CInt(Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)) Step 1
                                    Line Input #iArquivo, cLinhaArquivo
                                    If Mid(cLinhaArquivo, 1, 3) = "029" Then
                                        cLinha = cLinha & Mid(cLinhaArquivo, 12, Len(cLinhaArquivo) - 12) & vbCrLf
                                    End If
                                Next iVezes
                            End If

                        Case 30: 'Verifica se o campo é o 030 para mostrar a mensagem para o operador
                            If cLinha <> "" Then
                                'Mensagem  Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                            Else
                                If Dir(strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001") <> "" Then
                                    MataArquivo (strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001")
                                    'Mensagem  Mid(cLinhaArquivo, 11, Len(cLinhaArquivo) - 10)
                                    ImprimeGerencial = -1
                                End If
                            End If
                    End Select
                Wend
            End If
        Wend

        'Cria o arquivo temporário IMPRIME.TXT com a imagem do comprovante
        If (cLinha <> "") Then
            Close iArquivo
            CodigoFilial = 1
            Call GravaArquivo_Binario(App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT", cLinha)
            Exit For
        End If

        Sleep (1000)

        'O arquivo INTPOS.STS não retornou em 7 segundos, então o operador é informado.
        If (iTentativas = 7) Then

            MataArquivo (strCaminhoGerenciadorPadrao & "\REQ\INTPOS.001")
            'Mensagem "Gerenciador Padrão não está ativo!"
            ImprimeGerencial = 0
            Exit For
        End If
        If (ImprimeGerencial = 0) Or (ImprimeGerencial = -1) Then
            Close iArquivo
            Exit For
        End If
    Next iTentativas

    MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.STS")
    MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001")

    If Dir(App.Path + "\01\IMPRIME.TXT") <> "" Then
        'Bloqueia o teclado e o mouse para a impressão do TEF
        'iRetorno = Bematech_FI_IniciaModoTEF()

        ''''''''IMPRESSÃO DO RELATÓRIO GERENCIAL'''''''''''

        For via = 1 To 2 Step 1
            iArquivo = FreeFile
            Open App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT" For Input As iArquivo

            While Not EOF(iArquivo)
            '''''''''''''Lê uma linha do arquivo INTPOS.001 e grava em cLinhaArquivo
                Line Input #iArquivo, cLinhaArquivo
                'A função de impressão não aceita strings vazias
                If cLinhaArquivo = "" Then
                    cLinhaArquivo = " "
                End If

                '''''''''''''Imprime o que foi lido
                iRetorno = Bematech_FI_RelatorioGerencial(cLinhaArquivo & vbCrLf)

                '''''''''''''Aqui é feito o tratamento de erro de comunicação com a impressora
                '''''''''''''(desligamento da impressora durante a impressão do comprovante).
                If Not (VerificaRetornoFuncaoImpressora(iRetorno)) Then
                    'iRetorno = Bematech_FI_FinalizaModoTEF()
                    mdiPrincipal.MousePointer = vbDefault
                    If (MsgBox("A impressora não responde!" & vbCrLf & _
                        "Deseja imprimir novamente?", vbYesNo + vbQuestion, "Atenção") = vbYes) Then
                        Close iArquivo
                        iRetorno = Bematech_FI_FechaRelatorioGerencial
                        ImprimeGerencial (iConta)
                        Exit Function
                    Else
                        Close iArquivo
                        iRetorno = Bematech_FI_FechaRelatorioGerencial
                        ImprimeGerencial = 0
                        Exit Function
                    End If
                End If
            Wend



            '''''''''''''Aciona o corte de papel
            If via = 1 Then
                '''''''''''''Pula 7 linhas
                cSaltaLinha = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
                iRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(cSaltaLinha)
                iRetorno = Bematech_FI_VerificaTipoImpressora(tipoImpressora)
                If ((tipoImpressora = "2") Or (tipoImpressora = "4") Or (tipoImpressora = "6") Or (tipoImpressora = "8")) Then
                    'iRetorno = Bematech_FI_AcionaGuilhotinaMFD(0)
                End If
                '''''''''''''Exibe mensagem na tela
                'Mensagem "Por favor, destaque a " & via & "ª via."
                Sleep (3000)
            End If

            Close iArquivo
            'Mensagem ""
        Next via
        Close iArquivo
        iRetorno = Bematech_FI_FechaRelatorioGerencial()
        VerificaRetornoFuncaoImpressora (iRetorno)
    End If

    'Desbloqeia o teclado e o mouse
    iRetorno = Bematech_FI_FinalizaModoTEF()
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT")

End Function
''////////////////////////////////////////////////////////////////////////////////
''//
''// Função:
''//    VerificaRetornoFuncaoImpressora
''// Objetivo:
''//    Verificar o retorno da impressora e da função utilizada
''// Retorno:
''//    True para OK
''//    False para não OK
''//
''////////////////////////////////////////////////////////////////////////////////
Function VerificaRetornoFuncaoImpressora(ByVal iRetorno As Integer) As Boolean

   Dim cMSGErro As String
   Dim iACK As Integer
   Dim iST1 As Integer
   Dim iST2 As Integer

   'If strCaminhoGerenciadorPadrao = "" Then BuscaInformacoesEmpresa

   iACK = 0: iST1 = 0: iST2 = 0

    cMSGErro = ""
    VerificaRetornoFuncaoImpressora = False
    Select Case iRetorno
        Case 0:
           cMSGErro = "Erro de Comunicação !"
        Case -1:
            cMSGErro = "Erro de execução na Função !"
        Case -2:
            cMSGErro = "Parâmetro inválido na Função !"
        Case -3:
            cMSGErro = "Alíquota não Programada !"
        Case -4:
            cMSGErro = "Arquivo BEMAFI32.INI não Encontrado !"
        Case -5:
            cMSGErro = "Erro ao abrir a Porta de Comunicação !"
        Case -6:
            cMSGErro = "Impressora Desligada ou Cabo de Comunicação Desconectado !"
        Case -7:
            cMSGErro = "Código do Banco não encontrado no arquivo BEMAFI32.INI !"
        Case -8:
            cMSGErro = "Erro ao criar ou gravar arquivo STATUS.TXT ou RETORNO.TXT !"
        Case -27:
            cMSGErro = "Status diferente de 6, 0, 0 !"
        Case -30:
            cMSGErro = "Função incompatível com a impressora fiscal YANCO !"
    End Select

    If cMSGErro <> "" Then 'IF1
        Call Bematech_FI_FinalizaModoTEF
        VerificaRetornoFuncaoImpressora = False
    End If

    cMSGErro = ""
    If iRetorno = 1 Then 'IF2

        Call Bematech_FI_RetornoImpressora(iACK, iST1, iST2)
        If iACK = 21 Then 'IF2-1
            Call Bematech_FI_FinalizaModoTEF
            MsgBox "A Impressora retornou NAK !" & vbCrLf & _
                                       "Erro de Protocolo de Comunicação !", vbOKOnly, _
                                       "Atenção"
            VerificaRetornoFuncaoImpressora = False

        Else 'ELSEIF2-1
            If (iST1 <> 0) Or (iST2 <> 0) Then 'IF2-1-1
                  ' Analisa ST1
                If (iST1 >= 128) Then 'IF2-1-1-1
                    iST1 = iST1 - 128
                    cMSGErro = cMSGErro & "Fim de Papel" & vbCrLf
                End If 'ENDIF2-1-1-1
                If (iST1 >= 64) Then 'IF2-1-1-2
                    iST1 = iST1 - 64
                    cMSGErro = cMSGErro & "Pouco Papel" & vbCrLf
                    VerificaRetornoFuncaoImpressora = True
                    Exit Function
                End If 'ENDIF2-1-1-2
                If (iST1 >= 32) Then 'IF2-1-1-3
                    iST1 = iST1 - 32
                    cMSGErro = cMSGErro & "Erro no Relógio" & vbCrLf
                End If 'ENDIF2-1-1-3
                If (iST1 >= 16) Then 'IF2-1-1-4
                    iST1 = iST1 - 16
                    cMSGErro = cMSGErro & "Impressora em Erro" & vbCrLf
                End If 'ENDIF2-1-1-4
                If (iST1 >= 8) Then 'IF2-1-1-5
                    iST1 = iST1 - 8
                    cMSGErro = cMSGErro & "Primeiro Dado do Comando não foi ESC" & vbCrLf
                End If 'ENDIF2-1-1-5
                If iST1 >= 4 Then 'IF2-1-1-6
                    iST1 = iST1 - 4
                    cMSGErro = cMSGErro & "Comando Inexistente" & vbCrLf
                End If 'ENDIF2-1-1-6
                If iST1 >= 2 Then 'IF2-1-1-7
                    iST1 = iST1 - 2
                    cMSGErro = cMSGErro & "Cupom Fiscal Aberto" & vbCrLf
                End If 'ENDIF2-1-1-7
                If iST1 >= 1 Then 'IF2-1-1-8
                    iST1 = iST1 - 1
                    cMSGErro = cMSGErro & "Número de Parâmetros Inválidos" & vbCrLf
                End If 'ENDIF2-1-1-8
                'Analisa ST2
                If iST2 >= 128 Then 'IF2-1-1-9
                    iST2 = iST2 - 128
                    cMSGErro = cMSGErro & "Tipo de Parâmetro de Comando Inválido" & vbCrLf
                End If 'ENDIF2-1-1-9
                If iST2 >= 64 Then 'IF2-1-1-10
                    iST2 = iST2 - 64
                    cMSGErro = cMSGErro & "Memória Fiscal Lotada" & vbCrLf
                End If 'ENDIF2-1-1-10
                If iST2 >= 32 Then 'IF2-1-1-11
                    iST2 = iST2 - 32
                    cMSGErro = cMSGErro & "Erro na CMOS" & vbCrLf
                End If 'ENDIF2-1-1-11
                If iST2 >= 16 Then 'IF2-1-1-12
                    iST2 = iST2 - 16
                    cMSGErro = cMSGErro & "Alíquota não Programada" & vbCrLf
                End If 'ENDIF2-1-1-12
                If iST2 >= 8 Then 'IF2-1-1-13
                    iST2 = iST2 - 8
                    cMSGErro = cMSGErro & "Capacidade de Alíquota Programáveis Lotada" & vbCrLf
                End If 'ENDIF2-1-1-13
                If iST2 >= 4 Then 'IF2-1-1-14
                     iST2 = iST2 - 4
                     cMSGErro = cMSGErro & "Cancelamento não permitido" & vbCrLf
                End If 'ENDIF2-1-1-14
                If iST2 >= 2 Then 'IF2-1-1-15
                    iST2 = iST2 - 2
                    cMSGErro = cMSGErro & "CGC/IE do Proprietário não Programados" & vbCrLf
                End If 'ENDIF2-1-1-15
                If iST2 >= 1 Then 'IF2-1-1-16
                    iST2 = iST2 - 1
                    cMSGErro = cMSGErro & "Comando não executado" & vbCrLf
                End If 'ENDIF2-1-1-16
                If (cMSGErro <> "") Then 'IF2-1-1-17
                    Call Bematech_FI_FinalizaModoTEF
                    If cMSGErro <> "Comando não executado" & vbCrLf Then
                        MsgBox cMSGErro, vbOKOnly + vbExclamation, "Atenção"
                    End If
                    If VerificaRetornoFuncaoImpressora = True Then
                        VerificaRetornoFuncaoImpressora = False
                    End If
                End If 'ENDIF2-1-1-17
            Else
                VerificaRetornoFuncaoImpressora = True
            End If 'ENDIF2-1-1
        End If 'ENDIF2-1
    End If 'ENDIF2

End Function
'Public Sub CarregarFormasPagamento()
'    Dim formasPagto As New Collection
'    Dim formasdePagamento As String
'
'    Dim i As Long
'    Dim j As Integer
'    Dim tamanho As Integer
'    Dim Item As Variant
'
'    ' Verifica se existe o arquivo TEF.TXT, indicando que houve uma queda de
'    ' energia e que existe uma transação pendente.
'    formasdePagamento = Space(3016)
'    response = Bematech_FI_VerificaFormasPagamento(formasdePagamento)
'    j = 3016
'    Set formasPagto = Nothing
'    tamanho = 16
'    For i = 1 To j Step 58
'        formasPagto.Add (Mid(formasdePagamento, i, tamanho))
'    Next i
'    For Each Item In formasPagto
'        If Trim(Item) <> "" Then
'            ' frmTEFVariosCartoes.cboFormaPagto.AddItem (Trim(Item))
'        End If
'    Next Item
'
'End Sub
'Public Sub CancelarTransacoesPendentes()
'    Dim iArquivo As Integer
'    iArquivo = FreeFile
'    Open App.Path + "\TEF.TXT" For Input As iArquivo
'    'Lê o conteúdo do arquivo
'    If Not EOF(iArquivo) Then
'        Line Input #iArquivo, linhaArquivo
'    End If
'    Close iArquivo
'
'    'Se leu algo do arquivo então...
'    If linhaArquivo <> "" Then
'        For i = 0 To Len(linhaArquivo) Step 1
'            'Se o que leu for numérico...
'            If IsNumeric(Mid(linhaArquivo, i + 1, 1)) Then
'                'o auxiliar cLinha1 recebe o conteúdo numérico de cLinha
'                Call NaoConfirmaTransacao(CInt(Mid(linhaArquivo, i + 1, 1)))
'            End If
'        Next i
'    End If
'End Sub
Public Sub MataArquivo(ByVal caminho As String)
    If Dir(caminho) <> "" Then
            Kill caminho
    End If
End Sub
Public Sub GravaArquivo_Binario(ByVal caminho As String, ByVal dados As String)
    Dim iArquivo As Integer

    iArquivo = FreeFile
    Open caminho For Binary As iArquivo
        ' Escreve no arquivo
        Put iArquivo, , dados
        ' Fecha o arquivo
    Close iArquivo
End Sub
'Public Sub GravaArquivo_Random(ByVal caminho As String, ByVal dados As String)
'    Dim iArquivo As Integer
'
'    iArquivo = FreeFile
'    Open caminho For Random As iArquivo
'        ' Escreve no arquivo
'        Put #iArquivo, , dados
'        ' Fecha o arquivo
'    Close iArquivo
'
'End Sub
'
'
'Public Sub GravaArquivo_Output(ByVal caminho As String, ByVal dados As String)
'    Dim iArquivo As Integer
'
'    iArquivo = FreeFile
'    Open caminho For Output As iArquivo
'        ' Escreve no arquivo
'        Print #iArquivo, , dados
'        ' Fecha o arquivo
'    Close iArquivo
'
'End Sub
'
'
Public Sub FechamentoTEF(ValorPago As Currency, NUMEROCUPOM As String, hora As String)
On Error GoTo ERRO_TRATA
    Dim Gerencial As Boolean
    Dim iQuantasTransacoes As Integer
    Dim cValorPago As String

    'Mensagem ""
    cValorPago = ValorPago
    cValorPago = Replace(cValorPago, ",", "", , , vbTextCompare)

    'Mensagem "FechamentoTEF: Imprimindo Transação"

    Gerencial = False
    ''''''''''''IMPRIMINDO TRANSAÇÕES''''''''''''
    iQuantasTransacoes = 2
    iTransacao = 1
    If Not (ImprimeTransacao("CARTAO", cValorPago, NUMEROCUPOM, cHora, iTransacao, False)) Then

        'If MsgBox("A impressora não responde!" & vbCrLf & "Deseja imprimir novamente?", vbYesNo + vbInformation, "Atenção") = vbYes Then
            Gerencial = True


            ImprimeTransacao "CARTAO", cValorPago, NUMEROCUPOM, cHora, iTransacao, True
        'Else
            ''''''''''''SE OPTAR POR NÃO IMPRIMIR AS TRANSAÇÕES NOVAMENTE,
            ''''''''''''SERÁ FEITA A NÃO CONFIRMAÇÃO DELAS
        '    NaoConfirmaTransacao (iTransacao)
        'End If
    End If

    'Mensagem "FechamentoTEF: Confirmando Transação"

    ''''''''''''CONFIRMANDO A ÚLTIMA TRANSAÇÃO (EM CASO DE NÃO TER SIDO FEITA
    ''''''''''''A NÃO CONFIRMAÇÃO)
    If ((iQuantasTransacoes - 1) = iTransacao) Then
        ConfirmaTransacao (iTransacao)
    End If
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\PENDENTE.TXT")
    ''''''''''''MATANDO OS ARQUIVOS RESTANTES''''''''''''
    'For iVezes = 1 To (iQuantasTransacoes - 1) Step 1
        If Dir(strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001") <> "" Then
           MataArquivo (strCaminhoGerenciadorPadrao & "\RESP\INTPOS.001")
           MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\IMPRIME.TXT")
        End If
    'Next iVezes
    MataArquivo (App.Path & "\" & Format(CodigoFilial, "00") & "\TEF.TXT")

    Exit Sub
ERRO_TRATA:
    'ControleErros Err.Number, Err.Description, Err.Source, ""
End Sub

'Sub ModoAdministrativoTEF(CodigoFilial As Integer)

'    'Chama o módulo para limpar as variáveis para efetuar a venda com cartão.
'    Call TEFLimpaVariaveis
'
'    Call TEFModoAdministrativo(CodigoFilial)
'    If sTEFRetorno = "0" Then
'    ' ---> Neste ponto você vai chamar a sua rotina de impressão
'    '      que vai mover "0" para sTEFRetorno se a impressão foi concluída
'
'
'        ImprimeGerencial (1)
'
'        If sTEFRetorno = "0" Then
'            Call TEFFechaOperacao(CodigoFilial)
'        Else
'            Call TEFNaoConfirmaOperacao(CodigoFilial)
'        End If
'    End If
'
'    Exit Sub
'ERRO_TRATA:
'    If Err.Number <> 0 Then
'        ControleErros Err.Number, Err.Description, Err.Source, ""
'        Resume
'    End If
'End Sub
'
'
'
' ---< Cancela Transação efetivada >---
Public Sub TEFCancelaTransacao()
On Error GoTo ERRO_TRATA
    sTEFRetorno = "0"
    
    Call TEFVerificaGerenciadorAtivo
    If sTEFRetorno = "0" Then Call TEFGravaOperacao("CNC")
    If sTEFRetorno = "0" Then Call TEFConfirmaOperacao
    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFCancelaTransacao

' ---< Executa a confirmação da transação por meio de cartão >---
Public Sub TEFConfirmaOperacao()
On Error GoTo ERRO_TRATA
    Dim sLinha As String
    Dim iLinhas As Integer
  
    sTEFRetorno = "0"
  
    Call TEFVerificaTEFTmp("S")
    If sTEFRetorno = "0" Then
        iLinhas = 0
        Open (sTEFPath & "Resp\IntPos.001") For Input As #2
        Open (sTEFMsPath & "TEF.Imp") For Output As #4
        Do
            Line Input #2, sLinha
            If Left(sLinha, 3) = "029" Then
                Print #4, TEFRemoveAspas(RightFROMPos(sLinha, 11))
                iLinhas = iLinhas + 1
            End If
        Loop Until EOF(2)
        Close #2
        Close #4
        If iLinhas = 0 Then Kill (sTEFMsPath & "TEF.Imp")
        Call FileCopy(sTEFPath & "RESP\IntPos.001", sTEFPath & "RESP\IntPos" & CStr(iConta) & ".001")
    End If
    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub

Public Sub TEFVerificaTEFTmp(ByVal sExibeErro As String)
On Error GoTo ERRO_TRATA
    sTEFRetorno = "0"
    'Uma Forma de Realizar os testes de desligamento com + de dois Cartoes
    If Dir(sTEFPath & "Resp\IntPos5.001") <> "" Then
       Call FileCopy(sTEFPath & "Resp\IntPos2.001", sTEFPath & "RESP\IntPos.001")
       Open (sTEFPath & "Resp\IntPos.001") For Input As #2
       If EOF(2) Then sTEFRetorno = "1"
       Close #2
       iConta = 5
       sTEFRetorno = "0"
       GoTo Sai
    Else
       sTEFRetorno = "1"
    End If
    If Dir(sTEFPath & "Resp\IntPos4.001") <> "" Then
       Call FileCopy(sTEFPath & "Resp\IntPos2.001", sTEFPath & "RESP\IntPos.001")
       Open (sTEFPath & "Resp\IntPos.001") For Input As #2
       If EOF(2) Then sTEFRetorno = "1"
       Close #2
       iConta = 4
       sTEFRetorno = "0"
       GoTo Sai
    Else
       sTEFRetorno = "1"
    End If
    If Dir(sTEFPath & "Resp\IntPos3.001") <> "" Then
       Call FileCopy(sTEFPath & "Resp\IntPos2.001", sTEFPath & "RESP\IntPos.001")
       Open (sTEFPath & "Resp\IntPos.001") For Input As #2
       If EOF(2) Then sTEFRetorno = "1"
       Close #2
       iConta = 3
       sTEFRetorno = "0"
       GoTo Sai
    Else
       sTEFRetorno = "1"
    End If
    If Dir(sTEFPath & "Resp\IntPos2.001") <> "" Then
       Call FileCopy(sTEFPath & "Resp\IntPos2.001", sTEFPath & "RESP\IntPos.001")
       Open (sTEFPath & "Resp\IntPos.001") For Input As #2
       If EOF(2) Then sTEFRetorno = "1"
       Close #2
       iConta = 2
       sTEFRetorno = "0"
       GoTo Sai
    Else
       sTEFRetorno = "1"
    End If
    If Dir(sTEFPath & "Resp\IntPos1.001") <> "" Then
        Open (sTEFPath & "Resp\IntPos1.001") For Input As #2
        If EOF(2) Then sTEFRetorno = "1"
        Close #2
        sTEFRetorno = "0"
        DesligouImpressora = True
        GoTo Sai
    Else
        sTEFRetorno = "1"
    End If
    
    If Dir(sTEFPath & "Resp\IntPos.001") <> "" Then
        Open (sTEFPath & "Resp\IntPos.001") For Input As #2
        Call FileCopy(sTEFPath & "Resp\IntPos.001", sTEFPath & "RESP\IntPos" & iConta & ".001")
        If EOF(2) Then sTEFRetorno = "1"
        Close #2
        sTEFRetorno = "0"
        GoTo Sai
    Else
        sTEFRetorno = "1"
    End If
    
    'Neste Caso Vou Fazer busca se desligou o computador na impressao do cancelamento
    If Dir(sTEFPath & "Resp\CANCEL2.001") <> "" Then
       Call FileCopy(sTEFPath & "Resp\CANCEL2.001", sTEFPath & "RESP\IntPos.001")
       Call FileCopy(sTEFPath & "Resp\CANCEL2.001", sTEFPath & "RESP\IntPos3.001")
       Open (sTEFPath & "Resp\IntPos.001") For Input As #2
       If EOF(2) Then sTEFRetorno = "1"
       Close #2
       iConta = 3
       sTEFRetorno = "0"
       GoTo Sai
    Else
       sTEFRetorno = "1"
    End If
    If Dir(sTEFPath & "Resp\CANCEL1.001") <> "" Then
       Call FileCopy(sTEFPath & "Resp\CANCEL1.001", sTEFPath & "RESP\IntPos.001")
       Call FileCopy(sTEFPath & "Resp\CANCEL1.001", sTEFPath & "RESP\IntPos2.001")
       Open (sTEFPath & "Resp\IntPos.001") For Input As #2
       If EOF(2) Then sTEFRetorno = "1"
       Close #2
       iConta = 2
       sTEFRetorno = "0"
       GoTo Sai
    Else
       sTEFRetorno = "1"
    End If
    
    
    'Esse Teste e para quando o Computador for desligado ele exclui td da pasta c:\tef_dial e com isso tem que criar um
    'backup o arquivo intpos.001 para buscar as informacoes, emanoel
    
    If Dir(sTEFPath & "Back\IntPos1.001") <> "" Then
       Call FileCopy(sTEFPath & "Back\IntPos1.001", sTEFPath & "RESP\IntPos.001")
       Call FileCopy(sTEFPath & "Back\IntPos1.001", sTEFPath & "RESP\IntPos1.001")
       Kill (sTEFPath & "back\IntPos1.001")
       Open (sTEFPath & "RESP\IntPos.001") For Input As #2
       If EOF(2) Then sTEFRetorno = "1"
       Close #2
       sTEFRetorno = "0"
    Else
       sTEFRetorno = "1"
    End If
Sai:
    If (sTEFRetorno <> "0") And (sExibeErro = "S") Then MsgBox ("Não existe nenhuma operação pendente")

    
    Exit Sub
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFVerificaTEFTmp

' ---< Retorna os caracteres a direita de sString a partir da posição iPos >---
Public Function RightFROMPos(ByVal sString As String, ByVal iPos As Integer) As String
On Error GoTo ERRO_TRATA
    RightFROMPos = Trim(Mid(sString, iPos, (Len(sString) - iPos + 1)))
    
    Exit Function
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Function

' ---< Remove as aspas iniciais e finais de sString >---
Public Function TEFRemoveAspas(ByVal sString As String) As String
On Error GoTo ERRO_TRATA
    TEFRemoveAspas = Trim(Replace(sString, Chr(34), ""))
    
    Exit Function
ERRO_TRATA:
    If Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Function

' ---< Cria o arquivo IntPos.tmp para receber comandos >---
Public Sub TEFCriaArquivoREQIntPos001()
On Error GoTo ERRO_TRATA

    sTEFRetorno = "0"
    Open (sTEFPath & "REQ\IntPos.tmp") For Output As #3
    Close #3
    
    Exit Sub
ERRO_TRATA:
    If Err.Number = 76 Then
        
    ElseIf Err.Number <> 0 Then
        MsgBox "Erro nr: " & Err.Number & " Descrição: " & Err.Description & " Origem:" & Err.Source, 48, "TEF"
    End If
End Sub ' TEFCriaArquivoREQIntPos001

