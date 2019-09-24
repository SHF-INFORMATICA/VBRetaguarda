Attribute VB_Name = "mdlGERAL"
Public Sub ABRE_BANCO_SQLSERVER(Base_Dados As String)
'On Error GoTo ERRO_TRATA

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
   strFormatacaoKilo = "#####0.000"

Exit Sub
ERRO_TRATA:
   MsgBox "Erro ao abrir banco de dados : " & Err.Description
   End
End Sub

Public Sub ABRE_BANCO_GLOBAL()
'On Error GoTo ERRO_TRATA

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   If Not FSO.FileExists(App.Path & "\CONFIG.INI") Then
      'MsgBox "Arquivo de inicialização do sistema não encontrado, entre em contato com suporte."
      'End
      Exit Sub
   End If

   Dim Usuario_B        As String
   Dim Senha_B          As String
   Dim Nome_Banco       As String

   f = FreeFile

   Open App.Path & "\CONFIG.INI" For Input As f

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

   AUTENTICA_GRID_GLOBAL = "Provider=SQLOLEDB.1;Password='" & Senha_B & "'"
   AUTENTICA_GRID_GLOBAL = AUTENTICA_GRID_GLOBAL & ";Persist Security Info=True;User ID='" & Usuario_B & "'"
   AUTENTICA_GRID_GLOBAL = AUTENTICA_GRID_GLOBAL & ";Initial Catalog=" & Nome_Banco
   AUTENTICA_GRID_GLOBAL = AUTENTICA_GRID_GLOBAL & ";Data Source=" & Servidor_Global

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

Public Sub VERIFICA_SISTEMA()
'On Error GoTo ERRO_TRATA

   Dim Senha_Desbloqueio   As String
   Dim Dt_Liberacao_A      As String
   Dim Dt_Sistema_D        As Date

   Senha_Desbloqueio = ""
   Dt_Liberacao_A = ""
   Dt_Sistema_D = "01/01/1900"

   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

   SQL = "select par from EMPRESA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   tabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not tabEmpresa.EOF Then
      If Not IsNull(tabEmpresa.Fields("par").Value) Then
         Senha_Desbloqueio = "" & tabEmpresa.Fields("par").Value

         If Trim(Senha_Desbloqueio) = "" Then
            Senha_Desbloqueio = "" & InputBox("Atenção !!! Solicitar contra senha para liberação do sistema.", "Informe Chave para liberação.")
            GRAVA_PRIMEIRA_DATA Senha_Desbloqueio
            End
         End If

         frmCRIPTO.DECODIFICA (Senha_Desbloqueio)
         Dt_Liberacao_A = frmCRIPTO.txtDeCripto.Text

         If IsDate(Dt_Liberacao_A) Then _
            Dt_Sistema_D = Dt_Liberacao_A
         Else  'se for nulo entra aqui
            GRAVA_PRIMEIRA_DATA ""
            MsgBox "Solicitar Contra Senha ao suporte da SHF INFORMÁTICA."
            End
      End If
      Else
         GRAVA_PRIMEIRA_DATA ""
         MsgBox "Solicitar Contra Senha ao suporte da SHF INFORMÁTICA."
         End
   End If
   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

'============================
   If CDate(Date) = CDate(Dt_Sistema_D) Then
      Senha_Desbloqueio = "" & InputBox("Atenção !!! Solicitar contra senha para liberação do sistema.", "Informe Chave para liberação.")

      If Trim(Senha_Desbloqueio) <> "" Then
         frmCRIPTO.DECODIFICA (Senha_Desbloqueio)
         Dt_Liberacao_A = frmCRIPTO.txtDeCripto.Text

         If IsDate(Dt_Liberacao_A) Then _
            Dt_Sistema_D = Dt_Liberacao_A

         If CDate(Dt_Sistema_D) <= CDate(Date) Then
            MsgBox "Chave informada inválida, Solicitar Contra Senha ao suporte. " & Dt_Liberacao_A
            Else: GRAVA_PRIMEIRA_DATA Senha_Desbloqueio
         End If
      End If
      Else
         If CDate(Dt_Sistema_D) <= CDate(Date) Then
            Senha_Desbloqueio = "" & InputBox("Atenção !!! Solicitar contra senha para liberação do sistema." & Dt_Liberacao_A, "Informe Chave para liberação.")

            If Trim(Senha_Desbloqueio) = "" Then
               MsgBox "Chave informada inválida, Solicitar Contra Senha ao suporte da SHF INFORMÁTICA. " & Dt_Liberacao_A
               End
            End If

            frmCRIPTO.DECODIFICA (Senha_Desbloqueio)
            Dt_Liberacao_A = "" & frmCRIPTO.txtDeCripto.Text
            If IsDate(Dt_Liberacao_A) Then
               Dt_Sistema_D = Dt_Liberacao_A
               Else: Dt_Sistema_D = "01/01/1900"
            End If

            If Not IsDate(Dt_Sistema_D) Then
               MsgBox "Chave informada inválida, Solicitar Contra Senha ao suporte da SHF INFORMÁTICA. " & Dt_Liberacao_A
               End
            End If

            If Dt_Sistema_D <= Date Then
               MsgBox "Chave informada inválida, Solicitar Contra Senha ao suporte da SHF INFORMÁTICA. " & Dt_Liberacao_A
               End
               Else: GRAVA_PRIMEIRA_DATA Senha_Desbloqueio
            End If
            Else
               Dim DT_PEDIDO_ULTIMO As String
               Dim DT_INI
               Dim DT_FIM
               DT_PEDIDO_ULTIMO = "01/01/1900"

               If TabConsulta.State = 1 Then _
                  TabConsulta.Close

               SQL = "select max(dt_REQ) from PEDIDO WITH (NOLOCK)"
               TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not TabConsulta.EOF Then _
                  If Not IsNull(TabConsulta.Fields(0).Value) Then _
                     DT_PEDIDO_ULTIMO = TabConsulta.Fields(0).Value

               DT_INI = DMA(DT_PEDIDO_ULTIMO)
               DT_FIM = DMA(Date)

               If CDate(DT_INI) > CDate(DT_FIM) Then

                  MsgBox "Data do sistema incorreta, Atualizar sistema."

                  Senha_Desbloqueio = "" & InputBox("Solicitar contra senha para liberação do sistema.", "Informe Chave para liberação.")

                  frmCRIPTO.DECODIFICA (Senha_Desbloqueio)

                  SQL3 = frmCRIPTO.txtDeCripto.Text
                  DATA_INI = DMA(SQL3)

                  If Date >= DATA_INI Then
                     MsgBox "Chave informada inválida, Solicitar Contra Senha ao suporte da SHF INFORMÁTICA. " & Dt_Liberacao_A
                     Else: GRAVA_PRIMEIRA_DATA Senha_Desbloqueio
                  End If

                  GRAVA_PRIMEIRA_DATA Senha_Desbloqueio

                  End
               End If
         End If
   End If
'============================

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "VERIFICA_SISTEMA"
   End
End Sub

Public Sub INICIALIZA_SISTEMA()
'On Error GoTo ERRO_TRATA

   If Not FSO.FileExists(App.Path & "\MEGASIM.ini") Then
   'If Not FSO.FileExists("c:\megasim\MEGASIM.ini") Then
      MsgBox "Arquivo de inicialização do sistema não encontrado, entre em contato com suporte."
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
   ESTACAO_CPU = sLine

   Line Input #f, sLine
   NUMERO_CAIXA_CPU = sLine      'linha 6

   Line Input #f, sLine
   INDR_REMOTO = sLine           'linha 7

   Line Input #f, sLine
   USA_TEF = sLine               'linha 8

   Line Input #f, sLine
   USA_AUTTAR = sLine            'linha 9

   Line Input #f, sLine
   USA_POS = sLine               'linha 10

   Line Input #f, sLine
   INDR_TESTE = sLine            'linha 11

   Line Input #f, sLine
   USA_NFC_E = sLine             'linha 12

   Line Input #f, sLine
   INDR_SEQUENCIA = sLine        'linha 13

   Line Input #f, sLine
   INDR_OS_VEICULO = sLine       'linha 14 INDICA SE TRABALHA COMO OFICINA MECANICA

   Line Input #f, sLine
   'CaminhoNFE_A = sLine        'linha 15 CaminhoNFE_A

   Close #f

   If INDR_CAIXA = True Then
      If (App.PrevInstance) Then
          Dim nome_tela As String
          nome_tela = App.Title
          App.Title = "Sistama já está aberto, verifique !!!"
          'AppActivate  nome_tela
          SendKeys "%R", True
          MsgBox "Sistama já está aberto, verifique !!!"
          End
          Exit Sub
      End If
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "INICIALIZA_SISTEMA"
End Sub

Function TRAZ_NOME_MES(MES_ID_N As Integer) As String
'On Error GoTo ERRO_TRATA

   TRAZ_NOME_MES = ""

   If MES_ID_N = 1 Then
      TRAZ_NOME_MES = "Janeiro"
      Else
         If MES_ID_N = 2 Then
            TRAZ_NOME_MES = "Fevereiro"
            Else
               If MES_ID_N = 3 Then
                  TRAZ_NOME_MES = "Março"
                  Else
                     If MES_ID_N = 4 Then
                        TRAZ_NOME_MES = "Abril"
                        Else
                           If MES_ID_N = 5 Then
                              TRAZ_NOME_MES = "Maio"
                              Else
                                 If MES_ID_N = 6 Then
                                    TRAZ_NOME_MES = "Junho"
                                    Else
                                       If MES_ID_N = 7 Then
                                          TRAZ_NOME_MES = "Julho"
                                          Else
                                             If MES_ID_N = 8 Then
                                                TRAZ_NOME_MES = "Agosto"
                                                Else
                                                   If MES_ID_N = 9 Then
                                                      TRAZ_NOME_MES = "Setembro"
                                                      Else
                                                         If MES_ID_N = 10 Then
                                                            TRAZ_NOME_MES = "Outubro"
                                                            Else
                                                               If MES_ID_N = 11 Then
                                                                  TRAZ_NOME_MES = "Novembro"
                                                                  Else
                                                                     If MES_ID_N = 12 Then
                                                                        TRAZ_NOME_MES = "Desembro"
                                                                     End If
                                                               End If
                                                         End If
                                                   End If
                                             End If
                                       End If
                                 End If
                           End If
                     End If
               End If
         End If
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "TRAZ_NOME_MES"
End Function

Public Sub TRATA_ERROS(Desc_Erro As Variant, Formulario As String, Objeto As String)
   SQL3 = "Porfavor, descreva detalhes do(s) erro(s) : " & Err.Number & " - " & Desc_Erro
   Select Case Err.Number
      'Erros genéricos.
      Case 3005  'Database Name ' isn't a valid database name.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3006  'Database 'name' is exclusively locked.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3008  'Table 'name' is exclusively locked.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3009  'Couldn 't lock table 'name'; currently in use.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3010  'Table 'name' already exists.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3015  'Index name' isn't an index in this table.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3019  'Operation invalid without a current index.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3020  'Update or CancelUpdate without AddNew or Edit.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3021  'No current record.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3022  'Duplicate value in index, primary key, or relationship. Changes were unsuccessful.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3023  'AddNew or Edit already used.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3034  'Commit or Rollback without BeginTrans.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3036  'Database has reached maximum size.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3037  'can 't open any more tables or queries.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3040  'Disk I/O error during read.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3044  'Path' isn't a valid path.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3046  'Couldn 't save; currently locked by another user.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub

      'Erros relacionados a bloqueio de registros
      Case 3027  'can 't update.  Database or object is read-only.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3158  'Couldn 't save record; currently locked by another user.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3167  'Record is deleted.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3186  'Couldn 't save; currently locked by user 'name' on machine 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3187  'Couldn 't read; currently locked by user 'name' on machine 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3188  'Couldn 't update; currently locked by another session on this machine.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3189  'Table 'name' is exclusively locked by user 'name' on machine 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3197  'Data has changed; operation stopped.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3260  'Couldn 't update; currently locked by user 'name' on machine 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3261  'Table 'name' is exclusively locked by user 'name' on machine 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3356  'The database is opened by user 'name' on machine 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub

      'Erros relacionados a Permissões
      Case 3107  'Record(s) can't be added; no Insert Data permission on 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3108  'Record(s) can't be edited; no Update Data permission on 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3109  'Record(s) can't be deleted; no Delete Data permission on 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3110  'Couldn 't read definitions; no Read Definitions permission for table or query 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3111  'Couldn 't create; no Create permission for table or query 'name'.
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
      Case 3112  'Record(s) can't be read; no Read Data permission on 'name'.   End select
         CRITERIO_A = InputBox(SQL3, "Formulário: " & Formulario & "; Rotina:" & Objeto)
         GRAVA_ERRO Formulario, Objeto
         Exit Sub
   End Select

   CRITERIO_A = InputBox(SQL3 & " / " & Objeto, "Formulário: " & Formulario & "; Rotina:" & Objeto)
   
   GRAVA_ERRO Formulario, Objeto
   Exit Sub

'Resume Next Retorna a execução na linha que vem logo após à linha que gerou o erro
'Resume Executa mais uma vez a linha que gerou o erro.
'Resume Label Retorna a execução da linha que vem após a etiqueta citada.
'Resume Number Retorna a execução na linha com o número indicado.
'Exit Sub Sai da sub rotina atual.
'Exit Function Sai da função atual.
'Exit Property Sai da propriedade atual.
'On Error Redefine a lógica de tratamento de erros.
'Err.Clear Elimina o erro sem afetar a execução do programa.
'End Encerra a execução do programa.
'Number Fornece Numero do erro gerado
'Description Fornece a descrição do erro.
'Source Identifica o nome do objeto que gerou o erro
'Raise Gera um erro de execução, usado para testar condições de erro.
End Sub

Public Sub GRAVA_ERRO(Formulario As String, Objeto As String)
   If CONECTA_RETAGUARDA.State = 1 Then
      SqL2 = Replace(Err.Description, ",", " ")
      SqL2 = Replace(SqL2, "'", " ")

      CRITERIO_A = Replace(CRITERIO_A, ",", " ")
      CRITERIO_A = Replace(CRITERIO_A, "'", " ")

      SQL = "insert into ERRO values("
         SQL = SQL & MAX_ID("erro" & "_id", "ERRO", "", "", "", "")
         SQL = SQL & "," & Err.Number
         SQL = SQL & ",'" & Replace(SqL2, "',", " ") & "'"
         SQL = SQL & ",'" & CRITERIO_A & "'"
         SQL = SQL & ",'" & "FORMULÁRIO: " & Formulario & "   -   OBJETO: " & Objeto & "'"
         SQL = SQL & ",'" & Now & "'"
         SQL = SQL & "," & USUARIO_ID_N
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      CRITERIO_A = ""
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

Public Function CONFIRMA_PERGUNTA(Msg As String, Style As Variant, Title As String, Help As String, Ctxt As Long) As Boolean
'On Error GoTo ERRO_TRATA

   Dim Resposta_A As String

   CONFIRMA_PERGUNTA = False
   Beep

   Resposta_A = MsgBox(Msg, Style, Title, Help, Ctxt)
   If Trim(Resposta_A) = vbYes Then _
      CONFIRMA_PERGUNTA = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "CONFIRMA_PERGUNTA"
End Function


Public Function CALCULACNPJ(Numero As String) As String 'Funções para Validar CPF e CNPJ Validar CNPJ
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

   'Esta rotina foi adaptada da revista Fórum Access
   On Error GoTo Err_CPF

   Dim i As Integer        'utilizada nos FOR... NEXT
   Dim strCampo As String  'armazena do CPF que será utilizada para o cálculo
   Dim strCaracter As String   'armazena os dígitos do CPF da direita para a esquerda
   Dim intNumero As Integer    'armazena o digito separado para cálculo (uma a um)
   Dim intMais As Integer  'armazena o digito específico multiplicado pela sua base
   Dim lngSoma As Long     'armazena a soma dos dígitos multiplicados pela sua base(intmais)
   Dim dblDivisao As Double    'armazena a divisão dos dígitos * base por 11
   Dim lngInteiro As Long  'armazena inteiro da divisão
   Dim intResto As Integer     'armazena o resto
   Dim intDig1 As Integer  'armazena o 1º digito verificador
   Dim intDig2 As Integer  'armazena o 2º digito verificador
   Dim strConf As String   'armazena o digito verificador

   lngSoma = 0
   intNumero = 0
   intMais = 0
   strCampo = Left(CPF, 9)

   'Inicia cálculos do 1º dígito
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
   'Inicia cálculos do 2º dígito
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

'================================
Public Sub ParseToArray(sLine As String, a() As String)
'On Error GoTo ERRO_TRATA

   Dim p As Long, LastPos As Long, i As Long

   p = InStr(sLine, ";")

   Do While p

'Debug.Print Mid$(sLine, LastPos + 1, p - LastPos - 1)

On Error Resume Next
      a(i) = "" & Mid$(sLine, LastPos + 1, p - LastPos - 1)
      LastPos = p
      i = i + 1
      p = InStr(LastPos + 1, sLine, ";", vbBinaryCompare)
   Loop
   a(i) = Mid$(sLine, LastPos + 1)

Err.Clear

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
   MsgBox "O aplicativo não foi localizado !!! Verifique sua localização ...", vbExclamation
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

Public Function MAX_ID(Nome_Campo_Max As String, NOME_TABELA, Campo_01 As String, Info_01 As String, Campo_02 As String, Info_02 As String) As Long
'On Error GoTo ERRO_TRATA

   Dim TabID   As New ADODB.Recordset
   Dim strSQL  As String
 
   If TabID.State = 1 Then _
      TabID.Close
 
   MAX_ID = 0

   strSQL = "select max(" & Nome_Campo_Max & ") from " & NOME_TABELA & " "
   strSQL = strSQL & " where " & Nome_Campo_Max & " is not null "
   If (Trim(Campo_01) <> "") And (Info_01 <> "") Then _
      strSQL = strSQL & " and " & Trim(Campo_01) & " = " & Trim(Info_01)
   If (Trim(Campo_02) <> "") And (Info_02 <> "") Then _
      strSQL = strSQL & " and " & Trim(Campo_02) & " = " & Trim(Info_02)

   TabID.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
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
   ABRE_BANCO_SQLSERVER NOME_BANCO_DADOS
End Function
'============== por extenso
Public Function Extenso(nvalor)
'On Error GoTo ERRO_TRATA

    'Valida Argumento
   If IsNull(nvalor) Or nvalor <= 0 Or nvalor > 9999999.99 Then _
      Exit Function

    'Variáveis
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
      "milhões ", "milhão "), "")
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

Public Function INSERIR_0(TAMANHO_N As Long, STRING_A As String) As String
'On Error GoTo ERRO_TRATA

   INSERIR_0 = ""
   STRING_A = Trim(STRING_A)
   While Len(STRING_A) < TAMANHO_N
      STRING_A = "0" & STRING_A
      STRING_A = Trim(STRING_A)
   Wend

   INSERIR_0 = "" & Trim(STRING_A)

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "INSERIR_0"
End Function

Public Function INSERIR_BRANCO(TAMANHO_N As Long, STRING_A As String) As String
'On Error GoTo ERRO_TRATA

   INSERIR_BRANCO = ""
   STRING_A = Trim(STRING_A)
   While Len(STRING_A) < TAMANHO_N
      STRING_A = STRING_A & " "
   Wend

   INSERIR_BRANCO = STRING_A

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "INSERIR_BRANCO"
End Function

Public Sub DestacaTexto(Objeto As TextBox)
'On Error GoTo ERRO_TRATA

    Objeto.SelStart = 0
    Objeto.SelLength = Len(Objeto.Text)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "Modulo", "DestacaTexto"
End Sub

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
    'salvar as configurações atuais do form
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
    'restaura configurações do form
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

   If USUARIO_ID_N <> 144 Then
      Dim TabAcesso As New ADODB.Recordset

      If TabAcesso.State = 1 Then _
         TabAcesso.Close

      SQL = "select PERMISSAO.Menuid, PERMISSAO.Usuid, PERMISSAO.Acesso, USUARIO.USUARIO_ID, "
      SQL = SQL & " USUARIO.PESSOA_ID, USUARIO.EMPRESA_ID, USUARIO.NOME, USUARIO.SENHA, "
      SQL = SQL & " USUARIO.CPF, USUARIO.TIPO, "
      SQL = SQL & " USUARIO.Status, USUARIO.Logon, USUARIO.CLASSE"
      SQL = SQL & " from PERMISSAO WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN USUARIO WITH (NOLOCK)"
      SQL = SQL & " ON PERMISSAO.Usuid = USUARIO.USUARIO_ID"

      SQL = SQL & " where usuario_id = " & USUARIO_ID_N
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

Public Sub GERA_PEDIDO_ID_DEV()
'On Error GoTo ERRO_TRATA

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_pedido = seq_pedido + 1 "

   SQL = SQL & " from EMPRESA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   SQL = SQL & " AND cgc = '" & Trim(CNPJ_EMPRESA_N) & "'"

   CONECTA_RETAGUARDA.Execute SQL

   NUMR_PEDIDO_ID_N = 1

   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

   SQL = "select seq_pedido from EMPRESA WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK)"
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   tabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not tabEmpresa.EOF Then _
      If Not IsNull(tabEmpresa.Fields(0).Value) Then _
         NUMR_PEDIDO_ID_N = tabEmpresa.Fields(0).Value + 1
   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_PEDIDO_ID_DEV"
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

Public Function GERA_NUMR_PEDIDO_COMPRA() As Long
'On Error GoTo ERRO_TRATA

   GERA_NUMR_PEDIDO_COMPRA = 1

   SQL = "update EMPRESA set "
   SQL = SQL & " seq_pedcompra = seq_pedcompra + 1"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

   SQL = "select seq_pedcompra from EMPRESA WITH (NOLOCK)"
   SQL = SQL & " where empresa_id = " & EMPRESA_ID_N
   tabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not tabEmpresa.EOF Then _
      If Not IsNull(tabEmpresa.Fields(0).Value) Then _
         GERA_NUMR_PEDIDO_COMPRA = tabEmpresa.Fields(0).Value
   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_NUMR_PEDIDO_COMPRA"
End Function

'Funções para Validar CPF e CNPJ
'Validar CGC
Public Function VALIDA_CNPJCPF(CNPJ_CPF_A As String) As Boolean
'On Error GoTo ERRO_TRATA

   VALIDA_CNPJCPF = False

   If Len(CNPJ_CPF_A) <= 0 Then
      MsgBox "CNPJ/CPF inválido."
      Exit Function
   End If

   Select Case Len(CNPJ_CPF_A)
      Case Is = 11
        If Not CALCULACPF(CNPJ_CPF_A) Then
           'MsgBox "CPF com DV incorreto !!!  :  " & CNPJ_CPF_A
           Exit Function
        End If
      Case Is = 14
        If Not VALIDACGC(CNPJ_CPF_A) Then
           MsgBox "CNPJ com DV incorreto !!!  :  " & CNPJ_CPF_A
           Exit Function
        End If
      Case Is > 14
         MsgBox "CNPJ/CPF com DV incorreto !!!  :  " & CNPJ_CPF_A
         Exit Function
      Case Is < 11
         MsgBox "CNPJ/CPF com DV incorreto !!!  :  " & CNPJ_CPF_A
         Exit Function
   End Select

   VALIDA_CNPJCPF = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "VALIDA_CNPJCPF"
End Function

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

Public Sub BUSCA_ENDERECO_PESSOA(TIPO_A As String, CEP_A As String)

   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   SQL = "select ENDERECO_ID,PESSOA_ID,ENDERECO.CEP_ID,RUA,BAIRRO, "
   SQL = SQL & " COMPLEMENTO,TIPO,Numero,Cidade,UF,IBGE_ID"
   SQL = SQL & " from ENDERECO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
   SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"

   SQL = SQL & " where ENDERECO.pessoa_id = " & PESSOA_ID_N

   If Trim(TIPO_A) <> "" Then _
      SQL = SQL & " and tipo in ('" & Trim(TIPO_A) & "')"

   If Trim(CEP_A) <> "" Then _
      SQL = SQL & " and ENDERECO.CEP_ID = '" & Trim(CEP_A) & "'"

   tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub SP_PROC_BANCO(Codg_Banco As String)
   SQL = "select * from BANCO WITH (NOLOCK)"
   SQL = SQL & " where codg_banco = '" & Trim(Codg_Banco) & "'"
   TabBANCO.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub SP_GRAVA_CEP(CEP_A As String, CIDADE_A As String, UF_A As String, CODG_IBGE_N As Long)
   If TabCEP.State = 1 Then _
      TabCEP.Close

   If Trim(CEP_A) <> "" And Trim(CIDADE_A) <> "" And Trim(UF_A) <> "" Then
      SQL = "select * from CEP WITH (NOLOCK)"
      SQL = SQL & " where cep_ID = '" & Trim(CEP_A) & "'"
      SQL = SQL & " And cidade = '" & Trim(CIDADE_A) & "'"
      SQL = SQL & " And uf = '" & Trim(UF_A) & "'"
      TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabCEP.EOF Then
         SQL = "insert into CEP "
            SQL = SQL & " (CEP_id,Cidade,UF,IBGE_ID) "
         SQL = SQL & " values("
            SQL = SQL & "'" & Trim(CEP_A) & "'"
            SQL = SQL & ",'" & Trim(CIDADE_A) & "'"
            SQL = SQL & ",'" & Trim(UF_A) & "'"
            SQL = SQL & "," & Trim(CODG_IBGE_N)
         SQL = SQL & " )"
         Else
            SQL = "update CEP set"
               SQL = SQL & " cep_id = '" & Trim(CEP_A) & "'"
               SQL = SQL & ", Cidade = '" & Trim(CIDADE_A) & "'"
               SQL = SQL & ", UF = '" & Trim(UF_A) & "'"
               SQL = SQL & ", IBGE_ID = " & Trim(CODG_IBGE_N)
            SQL = SQL & " where cep_ID = '" & Trim(CEP_A) & "'"
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

   SQL = "select * from CEP WITH (NOLOCK)"
   SQL = SQL & " where cep_id = '" & Trim(CEP_A) & "'"
   TabCEP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub sp_Grava_Endereco(Cep As String, _
                             Rua As String, _
                             Bairro As String, _
                             Complemento As String, _
                             TIPO As String, _
                             Numero As String)

   Dim strSQL As String
   'SQL = "delete ENDERECO "
   'SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   'SQL = SQL & " and tipo = '" & Trim(TIPO) & "'"
   'CONECTA_RETAGUARDA.Execute SQL

   'Complemento = Replace(Complemento, "'", "´")
   'Complemento = Replace(Complemento, ",", ";")

   SP_MATA_ENDEREÇO TIPO

   BUSCA_ENDERECO_PESSOA TIPO, Cep

   If Not tabEndereco.EOF Then
      strSQL = "EXEC SP_UPDATE_ENDERECO '" & PESSOA_ID_N & "','" & Cep & "','" & Rua & "','" & Bairro & "','" & Complemento & "','" & TIPO & "','" & ENDERECO_ID_N & "','" & Numero & "'"
      Else
         ENDERECO_ID_N = MAX_ID("ENDERECO_ID", "endereco", "", "", "", "")
         strSQL = "EXEC SP_INSERT_ENDERECO '" & PESSOA_ID_N & "','" & Cep & "','" & Rua & "','" & Bairro & "','" & Complemento & "','" & TIPO & "'," & ENDERECO_ID_N & ",'" & Numero & "'"
   End If
   CONECTA_RETAGUARDA.Execute strSQL

   If tabEndereco.State = 1 Then _
      tabEndereco.Close
End Sub

Public Sub SP_MATA_ENDEREÇO(TIPO As String)

   SQL = "Delete IE "
   SQL = SQL & " where PESSOA_ID = " & PESSOA_ID_N
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "delete ENDERECO "
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and tipo = '" & Trim(TIPO) & "'"
   CONECTA_RETAGUARDA.Execute SQL

End Sub

Public Sub SP_PROCURA_PRODUTO(EMPRESA_ID As Long, _
                              Codg_Produto As String, _
                              FAMILIA_ID As Long, _
                              REFERENCIA As String, _
                              FORNEC_ID As Long, _
                              Codg_Barra As String, _
                              Tipo_Prod As Integer)

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select PRODUTO.PRODUTO_ID, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, "
   SQL = SQL & " PRODUTO.REFERENCIA, PRODUTO.FAMILIAPRODUTO_ID, PRODUTO.UNIDADE_MEDIDA, "
   SQL = SQL & " PRODUTO.SITUACAO, PRODUTO.TIPO_PROD, PRODUTO.FORNECEDOR_ID, PRODUTO.PRECO_CUSTO,"
   SQL = SQL & " PRODUTO.PRECO_ATACADO, PRODUTO.PRECO_Venda, PRODUTO.QTD_MINIMO, "
   SQL = SQL & " PRODUTO.QTD_MAXIMO, PRODUTO.PESO_LIQUIDO, PRODUTO.PESO_BRUTO,FAMILIAPRODUTO.CODG_FAMILIA,"
   SQL = SQL & " PRODUTO.PRODUTO_BALANCA, FAMILIAPRODUTO.DESCRICAO AS DescFamilia, FAMILIAPRODUTO.PRODUCAO "

   SQL = SQL & " from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN FAMILIAPRODUTO WITH (NOLOCK)"
   SQL = SQL & " ON PRODUTO.FAMILIAPRODUTO_ID = FAMILIAPRODUTO.FAMILIAPRODUTO_ID"

   SQL = SQL & " where empresa_id = " & EMPRESA_ID
   SQL = SQL & " and tipo_prod = " & Tipo_Prod

   If Trim(Codg_Produto) <> "" Then _
      SQL = SQL & " and codg_produto = '" & Trim(Codg_Produto) & "'"

   If FAMILIA_ID > 0 Then _
      SQL = SQL & " and familiaproduto_id = " & FAMILIA_ID

   If Trim(REFERENCIA) <> "" Then _
      SQL = SQL & " and referencia = '" & Trim(REFERENCIA) & "'"

   If FORNEC_ID > 0 Then _
      SQL = SQL & " and fornecedor_id = " & FORNEC_ID

   If Codg_Barra <> "" Then _
      SQL = SQL & " and CODG_BARRAS = '" & CODG_BARRAS & "'"

   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
End Sub

Public Sub CentralizaJanela2(Form As Form)
    Form.Top = (Screen.Height - Form.Height) / 2
    Form.Left = (Screen.Width - Form.Width) / 2
End Sub

Public Sub SP_PROCURA_FONE(NUMR_FONE As String)
   Dim strSQL As String

   If TabFone.State = 1 Then _
      TabFone.Close

   strSQL = "select * from FONE "
   strSQL = strSQL & " where pessoa_id = " & PESSOA_ID_N
   If Trim(NUMR_FONE) <> "" Then _
      strSQL = strSQL & " and numero = '" & NUMR_FONE & "'"
   TabFone.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
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
Public Function PreencheObjetos(adoControl As Object, strSQL As String, Grade As Variant) As String
On Error Resume Next

    adoControl.ConnectionString = strConexao
    adoControl.UserName = uid
    adoControl.Password = pwd
    adoControl.RecordSource = strSQL
    adoControl.CommandTimeout = 200
    Grade.Refresh
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
   Rs.Open "select getdate() as data", CONECTA_RETAGUARDA, , adCmdText
   If (DateValue(Rs("data")) <= DateValue(dataf)) Then
      data_atual = Rs("data")
   Else
      data_atual = dataf
   End If
   Rs.Close

   ultimo_dia_mes = "01/" & Format(DateAdd("m", 1, data_atual), "mm/yyyy")

   Rs.Open "select data from feriado WHERE data >= '" & DMA(datai) & "' and data < '" & DMA(ultimo_dia_mes) & "'", CONECTA_RETAGUARDA, , adCmdText
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
      MsgBox "Período incorreto...", vbCritical, "Atenção !!."
   End If
End Sub

Public Function Valida_Inscricao_Estadual(INSC_A As String, UF_A As String) As Integer
'On Error GoTo ERRO_TRATA

   Valida_Inscricao_Estadual = 0
   If Trim(UCase(INSC_A)) = "ISENTA" Then _
      INSC_A = "ISENTO"
   If Trim(INSC_A) = "" Then _
      INSC_A = "ISENTO"

   strRetorno = INSC_A
   INSC_A = Trim(INSC_A)
   INSC_A = Replace(INSC_A, ".", "")
   INSC_A = Replace(INSC_A, ",", "")
   INSC_A = Replace(INSC_A, "-", "")
   INSC_A = Replace(INSC_A, "/", "")

   Valida_Inscricao_Estadual = Inscricao(Trim(INSC_A), Trim(UF_A))

   If Valida_Inscricao_Estadual = 1 Then
       MsgBox "Inscrição Estadual inválida para a UF Informada . " & strRetorno, vbExclamation, "MEGASIM"
   ElseIf Valida_Inscricao_Estadual = 2 Then
       MsgBox "Cliente com parametros de Inscrição Estadual inválidos ." & strRetorno, vbExclamation, "MEGASIM"
   End If

Exit Function
ERRO_TRATA:
   MsgBox Err.Description
End Function

'FUNÇÃO QUE RETORNA O NUMERO DE DIAS NO MÊS
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
               Case "Â", "Ã", "Á", "À", "Ä", "Å"
                  C = C & "A"
               Case "É", "È", "Ê", "Ë"
                  C = C & "E"
               Case "Í", "Ì", "Î"
                  C = C & "I"
               Case "Ó", "Ò", "Ô", "Ö", "Õ"
                  C = C & "O"
               Case "Ú", "Ù", "Û", "Ü"
                  C = C & "U"
               Case "Ç"
                  C = C & "C"
               Case "Ñ"
                  C = C & "N"
               Case "Ÿ", "Ý"
                  C = C & "Y"
               Case "´", "`", "'"
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
   Grid.Override.CellClickAction = ssClickActionRowselect
   Grid.Override.selectedCellAppearance.BackColorAlpha = ssAlphaUseAlphaLevel

   Grid.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
End Sub

Public Function MontaSQLRelatorio(lngCodigoRelatorio As Long, Botao As Object) As String
'On Error GoTo ERRO_TRATA

   Dim rsBotao As New ADODB.Recordset
   
   Dim strFiltroRelatorio As String
   
   rsBotao.Open "select * from RelatorioFiltro WITH (NOLOCK) WHERE Codigo = " & lngCodigoRelatorio & " AND CodigoUsuario = " & intUsuario, CONECTA_RETAGUARDA, , adCmdText
   If Not rsBotao.EOF Then
       strFiltroRelatorio = Trim(rsBotao!selectionFormula & "")
   Else
       MontaSQLRelatorio = ""
       'MsgBox "Consulta inexistente...", 48, "Atenção..."
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
    'ControleErros Err.Number, Err.Description, Err.Source, "Monta botõa verde e vermelho..."
End Function

Public Function Relatorio(strFormulario As String, strRelatorio As String, strNomeRelatorio As String, strFiltro As String, Optional strParametro1 As String, Optional strParametro2 As String, Optional strParametro3 As String, Optional strParametro4 As String, Optional strParametro5 As String, Optional strParametro6 As String, Optional strParametro7 As String, Optional strParametro8 As String, Optional strParametro9 As String, Optional strParametro10 As String, Optional strParametro11 As String, Optional strParametro12 As String, Optional strParametro13 As String, Optional strParametro14 As String, Optional strParametro15 As String, Optional strParametro16 As String, Optional strParametro17 As String, Optional strParametro18 As String, Optional strParametro19 As String, Optional strParametro20 As String, Optional strParametro21 As String, Optional strParametro22 As String, Optional strParametro23 As String)
'On Error GoTo ERRO_TRATA

   Dim LS_Report As String
   Dim gsNomeRel As String
   Dim sParametro As String
   Dim sselectionFormula As String
   Dim strImpressora As String
   
   sselectionFormula = CRITERIO_A
   
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

   crxReport.RecordSelectionFormula = sselectionFormula
   crxReport.GroupSelectionFormula = selectionFormulaGrupo

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
   
    LogRelatorio strFormulario, LS_Report, Left(gsNomeRel, 200), Left(TrocaApostrofeSharpe(strParametro1), 200), Left(TrocaApostrofeSharpe(sselectionFormula), 200)
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
            'tava dando pau aq, aí coloquei o 1 na frente = bica
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
        crxReport.PrinterSetup frmINICIO.hwnd
        crxReport.PrintOut False, 1, True
        
        If MsgBox("Confirma Impressão?", vbQuestion + vbYesNo, "Gera Faturamento") = vbNo Then 'Wanderson - 25-02-2010
            Exit Function
        End If
        
'        If crxReport.DriverName <> "" Then
'            Printer.Orientation = crxReport.PaperOrientation
'            crxReport.DisplayProgressDialog = True
'            crxReport.selectPrinter crxReport.DriverName, crxReport.PrinterName, crxReport.PortName
'            crxReport.PaperOrientation = Printer.Orientation
'            crxReport.PrintOut True, 1
'        End If
    End If

    Exit Function
ERRO_TRATA:
    If Err.Number = "-2147189423" Then
        MsgBox "Atenção. Selecione uma impressora padrao."
        Exit Function
    ElseIf Err.Number = "-2147206461" Then
        MsgBox "Atenção. Relatório nao localizado Caminho: " & LS_Report
        Exit Function
    End If
'    ControleErros Err.Number, Err.Description, Err.Source, "GeraRelatorio"
End Function

'##ModelId=417E3D7402F4
Public Sub LS_Envia_Formula(LS_Formula1 As String, LS_Parametro1 As String)
    'configurar fórmula
    Set OG_Formula_Field = crxReport.FormulaFields
    For Each OG_Formula_Field In OG_Formula_Field
        If Trim(UCase(OG_Formula_Field.Name)) = Trim(UCase(LS_Formula1)) Then
            OG_Formula_Field.Text = LS_Parametro1
            Exit Sub
        End If
    Next
End Sub

Public Sub LS_Envia_FormulaSubReport(LS_Formula1 As String, LS_Parametro1 As String)
    'configurar fórmula
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
      If crxReport.ParameterFields.item(i).Name = LS_Formula1 And crxReport.ParameterFields.item(i).ValueType = 7 Then
         crxReport.ParameterFields.item(i).AddCurrentValue CDbl(LS_Parametro1)
         Exit Sub
      End If
      'valor
      If crxReport.ParameterFields.item(i).Name = LS_Formula1 And crxReport.ParameterFields.item(i).ValueType = 8 Then
         crxReport.ParameterFields.item(i).AddCurrentValue CDbl(LS_Parametro1)
         Exit Sub
      End If
      'texto
      If crxReport.ParameterFields.item(i).Name = LS_Formula1 And crxReport.ParameterFields.item(i).ValueType = 12 Then
         crxReport.ParameterFields.item(i).AddCurrentValue (LS_Parametro1)
         Exit Sub
      End If
      'date
      If crxReport.ParameterFields.item(i).Name = LS_Formula1 And (crxReport.ParameterFields.item(i).ValueType = 16 Or crxReport.ParameterFields.item(i).ValueType = 10) Then
         crxReport.ParameterFields.item(i).AddCurrentValue (DateValue(LS_Parametro1))
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
   'Note: CopyfromRecordset copies only the data and not the field
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
   PB.Max = Rs.RecordCount
   PB.Value = 0
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
      PB.Value = i
      k = 1
       For Each f In Rs.Fields
           tb.Cell(i + 1, k).Range.Text = Rs.Fields(f.Name)
            k = k + 1
       Next
       Rs.MoveNext
   Next

   Screen.MousePointer = vbNormal
   MsgBox "Conversão realizada com sucesso"
   PB.Value = 0

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
Public Function HabilitaCampos(Formulario As Form, VALOR As Boolean, Optional NHabilita As Boolean)
    On Error Resume Next
    Dim i As Long
    For i = 1 To Formulario.Count
        If Formulario.Controls(i).Tag <> "N" Then
            Formulario.Controls(i).Enabled = Not VALOR
        ElseIf Formulario.Controls(i).Tag = "N" And NHabilita = True Then
            Formulario.Controls(i).Enabled = VALOR
        End If
    Next i
End Function

'##ModelId=417E3D74020A
Public Sub CentraNaTela(f As Form)
    With frmINICIO
        If f.WindowState = vbNormal Then              'se o form não está minimizado
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
    ' Função      : Serve para formatar e retornar o mes dia e ano em formato Americano, e se passar o parametro retorna a hora
    ' Autor       :
    ' Alterações  :

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
    ' Função      : Converte um valor em Moeda
    ' Autor       :
    ' Alterações  :
    If Not IsNumeric(dado) Then dado = 0
        
    If CasasDecimais = 0 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0")
    ElseIf CasasDecimais = 1 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.0")
    ElseIf CasasDecimais = 2 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "Standard")
    ElseIf CasasDecimais = 3 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.000")
    ElseIf CasasDecimais = 4 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.0000")
    ElseIf CasasDecimais = 5 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.00000")
    ElseIf CasasDecimais = 6 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.000000")
    ElseIf CasasDecimais = 7 Then
        moeda = Format(CCur(IIf(IsNull(dado), 0, dado)), "###,##0.0000000")
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
Public Function TrocaApostrofeSharpe(VALOR As String) As String
   For i = 1 To Len(VALOR)
      If Mid(VALOR, i, 1) <> "'" Then
         TrocaApostrofeSharpe = TrocaApostrofeSharpe & Mid(VALOR, i, 1)
         Else: TrocaApostrofeSharpe = TrocaApostrofeSharpe & "'"
      End If
      DoEvents
   Next i
End Function

Public Sub LogRelatorio(Formulario As String, TipoRelatorio As String, NomeRpt As String, Periodo As String, selectionFormula As String)
    Dim varsql As String
    varsql = "INSERT INTO LogRelatorio (CodigoEmpresa, DataHora, CodigoUsuario, Formulario, TipoRelatorio, NomeRpt, Periodo, selectionFormula)"
    varsql = varsql & " VALUES ("
    varsql = varsql & intEmpresa
    varsql = varsql & ",getdate()" 'MARCADO PARA NAO MUDAR
    varsql = varsql & " , " & intUsuario
    varsql = varsql & " , '" & Formulario & "'"
    varsql = varsql & " , '" & TipoRelatorio & "'"
    varsql = varsql & " , '" & NomeRpt & "'"
    varsql = varsql & " , '" & Periodo & "'"
    varsql = varsql & " , '" & selectionFormula & "'"
    varsql = varsql & ")"
    CONECTA_RETAGUARDA.Execute varsql
End Sub

'##ModelId=417E3D74010E
Public Function ValidaCPF(NUM_CPF As String) As Boolean
    On Error Resume Next
    '//------------------------------------
    '// Função que testa se o CPF é válido
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

    'Selecionando apenas os números da stringue NUM_CPF
    For i = 1 To Len(Trim$(NUM_CPF))
        If Mid$(NUM_CPF, i, 1) <> "." And Mid$(NUM_CPF, i, 1) <> "-" Then
            CPF = CPF & Mid$(NUM_CPF, i, 1)
        End If
    Next i

    'Se o tamanho do CPF passado for < 11 então de imediato já é falso
    If Len(CPF) <> 11 Then
        ValidaCPF = False
        Exit Function
    End If

    aux = Right(CPF, 2) 'recebe os dois digitos para ser comparado

    CPF = Left(CPF, 9)

    'Cálculo do primeiro digito
    Calculo = 0
    Calculo = (Mid$(CPF, 9, 1) * 9) + (Mid$(CPF, 8, 1) * 8) + (Mid$(CPF, 7, 1) * 7) + (Mid$(CPF, 6, 1) * 6) + (Mid$(CPF, 5, 1) * 5) + (Mid$(CPF, 4, 1) * 4) + (Mid$(CPF, 3, 1) * 3) + (Mid$(CPF, 2, 1) * 2) + (Mid$(CPF, 1, 1) * 1)
    resto = (Calculo Mod 11)

    If resto = 10 Then
        resto = 0
    End If

    digitos = resto

    'Cálculo do segundo digito
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
   ' Dim rs As New ADODB.Recordset               'HOLDS ALL DATA RETURNED from QUERY

    'rs.CursorLocation = adUseClient
    'rs.Open SQL, MyDatabase, adOpenForwardOnly, adLockReadOnly, adCmdText
    'Set rs.ActiveConnection = Nothing
   '
    'If rs.EOF Then
    '    MsgBox "Não há registros!", vbInformation
    '    Unload Me
    '    Exit Sub
    'End If
          
    Dim crystal As New CRAXDRT.Application      'LOADS REPORT from FILE
    Dim report As CRAXDRT.report            'HOLDS REPORT
    Set report = crystal.OpenReport(PATH_REL & "RelVendasDiarias.rpt")  'OPEN OUR REPORT

    report.DiscardSavedData                      'CLEARS REPORT SO WE WORK from RECORDSET
    'report.Database.SetDataSource rs             'LINK REPORT TO RECORDSET
    report.RecordSelectionFormula = CRITERIO_A
    'report.GroupselectionFormula = selectionFormulaGrupo
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

Public Function BuscaCodigo(SQL As String, Banco As Variant, CAMPO As Variant) As String
'On Error GoTo ERRO_TRATA

    Dim rs_Maior As New ADODB.Recordset
    
    BuscaCodigo = ""
    rs_Maior.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
    If Not rs_Maior.EOF Then
        BuscaCodigo = IIf(Not IsNull(rs_Maior(CAMPO)), rs_Maior(CAMPO), 0)
    Else
        BuscaCodigo = 0
    End If
    rs_Maior.Close
    
    Exit Function
ERRO_TRATA:
'    ControleErros Err.Number, Err.Description, Err.Source, "codigoCombo"
End Function

Public Sub BUSCA_ALIQUOTA_ICMS(UF_ORIGEM_A As String, UF_DESTINO_A As String, CFOP_ID_N As Integer)
'On Error GoTo ERRO_TRATA

   Dim TabAliquota   As New ADODB.Recordset
   Dim strSQL        As String

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0

   ALIQUTOA_PIS_N = 0
   ALIQUTOA_COFINS_N = 0
   CST_PIS_A = ""
   CST_COFINS_A = ""
   CST_ICMS_A = ""

   If TabAliquota.State = 1 Then _
      TabAliquota.Close

   strSQL = "select ALIQUOTA_ICMS_DENTRO, ALIQUOTA_ICMS_FORA from ALIQUOTA_UF WITH (NOLOCK)"
   strSQL = strSQL & " INNER JOIN CFOPUF WITH (NOLOCK)"
   strSQL = strSQL & " ON ALIQUOTA_UF.CFOPUF_ID = CFOPUF.CFOPUF_ID"
   strSQL = strSQL & " INNER JOIN CFOP WITH (NOLOCK)"
   strSQL = strSQL & " ON CFOPUF.CFOP_ID = CFOP.CFOP_ID"

   strSQL = strSQL & " where uf_origem = '" & Trim(UF_ORIGEM_A) & "'"

   If Trim(UF_DESTINO_A) <> "" Then _
      strSQL = strSQL & " and uf_destino = '" & Trim(UF_DESTINO_A) & "'"

   If CFOP_ID_N > 0 Then _
      strSQL = strSQL & " and CFOP.cfop_id = " & CFOP_ID_N

   TabAliquota.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabAliquota.EOF Then
      ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0 & TabAliquota.Fields("aliquota_icms_dentro").Value
      ALIQUOTA_ICMS_NORMAL_FORA_UF = 0 & TabAliquota.Fields("aliquota_icms_fora").Value
   End If
   If TabAliquota.State = 1 Then _
      TabAliquota.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "BUSCA_ALIQUOTA_ICMS"
End Sub

Public Sub BUSCA_ALIQUOTA_PISCOFINS(UF_ORIGEM_A As String, UF_DESTINO_A As String, CFOP_ID_N As Integer)
'On Error GoTo ERRO_TRATA

   Dim TabAliquota   As New ADODB.Recordset
   Dim strSQL        As String

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0
PERC_BASE_REDUZ_N = 0
   ALIQUTOA_PIS_N = 0
   ALIQUTOA_COFINS_N = 0
   CST_PIS_A = ""
   CST_COFINS_A = ""
   CST_ICMS_A = ""

   If TabAliquota.State = 1 Then _
      TabAliquota.Close

   strSQL = "select ALIQUOTA_PIS, ALIQUOTA_COFINS,CST_PIS,CST_COFINS,CST_ICMS,PERC_BASE_REDUZ,CST_ORIG_ICMS"
   strSQL = strSQL & " from ALIQUOTA_UF WITH (NOLOCK)"
   strSQL = strSQL & " INNER JOIN CFOPUF WITH (NOLOCK)"
   strSQL = strSQL & " ON ALIQUOTA_UF.CFOPUF_ID = CFOPUF.CFOPUF_ID"
   strSQL = strSQL & " INNER JOIN CFOP WITH (NOLOCK)"
   strSQL = strSQL & " ON CFOPUF.CFOP_ID = CFOP.CFOP_ID"

   strSQL = strSQL & " where uf_origem = '" & Trim(UF_ORIGEM_A) & "'"
   'strSQL = strSQL & " and estabelecimento_ID = " & ESTABELECIMENTO_ID_N

   If Trim(UF_DESTINO_A) <> "" Then _
      strSQL = strSQL & " and uf_destino = '" & Trim(UF_DESTINO_A) & "'"

   If CFOP_ID_N > 0 Then _
      strSQL = strSQL & " and CFOP.cfop_id = " & CFOP_ID_N

   TabAliquota.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabAliquota.EOF Then
      ALIQUTOA_PIS_N = 0 & TabAliquota.Fields("ALIQUOTA_PIS").Value
      ALIQUTOA_COFINS_N = 0 & TabAliquota.Fields("ALIQUOTA_COFINS").Value

      CST_PIS_A = "" & TabAliquota.Fields("CST_PIS").Value
      CST_COFINS_A = "" & TabAliquota.Fields("CST_COFINS").Value
      CST_ICMS_A = "" & TabAliquota.Fields("CST_ICMS").Value

      PERC_BASE_REDUZ_N = 0 & TabAliquota.Fields("PERC_BASE_REDUZ").Value
      CST_ORIG_ICMS_N = 0 & TabAliquota.Fields("CST_ORIG_ICMS").Value
   End If
   If TabAliquota.State = 1 Then _
      TabAliquota.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "BUSCA_ALIQUOTA_PISCOFINS"
End Sub

Public Function fnAlinhaCampos(VALOR As String, tamanho As Integer, Ajuste As String, Caracter As String)
   Dim Concatena  As String
   Dim GrupoCaracter As String
   GrupoCaracter = ""
   
   For i = 1 To tamanho
       GrupoCaracter = GrupoCaracter & Caracter
   Next i


   Concatena = ""
   If Ajuste = "E" Then
       Concatena = Format(VALOR * 100, GrupoCaracter)
       'Concatena = "" 'Replace(Caracter, Tamanho - Len(Valor)) + Trim(Valor)
   Else
       Concatena = Format(VALOR * 100, GrupoCaracter)
       'Concatena = "" ''RTrim(Valor) + Replace(Caracter, Tamanho - Len(Valor))"
   End If
   If Len(VALOR) > tamanho Then Concatena = Left(Trim(VALOR), tamanho)
   fnAlinhaCampos = Concatena
End Function

Public Function SomenteNumeros(Enter As Integer) As Integer
    If Enter = 8 Or Enter = 13 Or Enter = 45 Then 'Se for DELETE ou Backspace não fazer nada
        SomenteNumeros = Enter
        Exit Function
    End If
    
    If Enter = 46 Or Enter = 44 Then  'Se for ponto ou vírcula assume vírgula
        SomenteNumeros = 44
        Exit Function
    End If
    
    If Not IsNumeric(Chr(Enter)) Then  'Se não for número retorne zero
        SomenteNumeros = 0
        Exit Function
    End If
    
    If Enter = 8 Then                 'Se for Delete ou Backspace não fazer nada
        SomenteNumeros = Enter
        Exit Function
    End If
    
    If Enter < 48 Or Enter > 57 Then  'Se não for número retorne zero
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

Public Function BUSCA_TRIBUTACAO_PRODUTO(ORIGEM_MERCADORIA_A As String, CST_A As String) As String
'On Error GoTo ERRO_TRATA

'CRT – Código do Regime Tributário 1,2 ou 3 ONDE:
   '1 = Simples Nacional
   '2 = Simples Nacional-Excesso de sublimite receita bruta
   '3 = Regime Normal - RPA

'CST_A = Código da Situação Tributária do produto ONDE:
   '00 Tributada integralmente Tributada integralmente
   '10 Tributada  e com cobrança do ICMS por substituição tributária  Tributada  e com cobrança do ICMS por substituição tributária
   '20 Com redução de base de cálculo   Com redução de base de cálculo
   '30 Isenta ou não tributada e com cobrança do ICMS por substituição tributária Isenta ou não tributada e com cobrança do ICMS por substituição tributária
   '40 Isenta Isenta
   '41 Não tributada  Não tributada
   '50 Suspensão Suspensão
   '51 Diferimento Diferimento
   '60 ICMS cobrado anteriormente por substituição tributária   ICMS cobrado anteriormente por substituição tributária
   '70 Com redução de base de cálculo e cobrança de ICMS por substituição tributária Com redução de base de cálculo e cobrança de ICMS por substituição tributária
   '90 Outras Outras

'CSOSN EMPRESAS NO REGIME TRIBUTÁRIO SIMPLES NACIONAL - Código de Situação da Operação no Simples Nacional ONDE:
   '101 Tributada pelo Simples Nacional com permissão de crédito                                                              classificam-se neste código as operações que permitem a indicação da alíquota de ICMS devido no Simples Nacional e o valor do crédito correspondente
   '102 Tributada pelo Simples Nacional sem permissão de crédito                                                              classificam-se código as operações que não permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor do crédito, e não estejam abrangidas nas hipóteses dos códigos 103, 203, 300, 400, 500 e 900
   '103 Isenção do ICMS no Simples Nacional para faixa de receita bruta                                                       classificam-se neste código as operações praticadas por optantes do Simples Nacional contempladas com isenção concedida para faixa de receita bruta nos termos da Lei Complementar n. 123 de 2006
   '201 Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por substituição tributária           classificam-se neste código as operações  que permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor crédito e com cobrança do ICMS por substituição tributária
   '202 Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por substituição tributária           classificam-se neste código as operações  que não permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor crédito, e não estejam abrangidas nas hipóteses dos códigos 103, 203, 300, 400, 500 e 900 e com cobrança do ICMS por substituição tributária
   '203 Isenção do ICMS no Simples Nacional para a faixa de receita bruta e com cobrança de ICMS por substituição tributária  classificam-se neste código as operações que praticadas por optantes do Simples Nacional contemplados com isenção para a faixa de receita bruta, mas com ICMS cobrado por substituição tributária
   '300 Imune                                                                                                                 classificam-se neste código as operações que praticadas por optantes do Simples Nacional contempladas com imunidade do ICMS
   '400 Não tributada pelo Simples Nacional                                                                                   classificam-se neste código as operações que praticadas por optantes do Simples NacionaL não sujeitas à tributação pelo ICMS dentro do Simples Nacional
   '500 ICMS cobrado anteriormente por substituição tributária                                                                classificam-se neste código as operações sujeitas exclusivamente ao regime de substituição tributária na condição de substituído tributário ou no caso de antecipações
   '900 Outros                                                                                                                classificam-se neste código as operações que não se enquadrem nos códigos 101, 102, 103, 201, 202, 203, 300, 400 e 500

   If Trim(ORIGEM_MERCADORIA_A) = "" Or Trim(CST_A) = "" Then _
      Exit Function

   Dim CODG_INTERNO_TRIBUTACAO As String

   CODG_INTERNO_TRIBUTACAO = "" & Trim(ORIGEM_MERCADORIA_A) & Trim(CST_A)

   BUSCA_TRIBUTACAO_PRODUTO = 400  'Não tributada pelo Simples Nacional

   Select Case CTR_EMPRESA_N  'REGIME TRIBUTARIO QUE A EMPRESA SE ENQUADRA
      Case 1   '1 = Simples Nacional
         'Tributada integralmente = 00 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "000" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 101

         'Tributada e com cobrança do ICMS por substituição tributária = 10 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "010" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 201

         'Com redução de base de cálculo = 20 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "020" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 101

         'Isenta ou não tributada e com cobrança do ICMS por substituição tributária = 30 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "030" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 203

         'Isenta = 40 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "040" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 400

         'Não tributada = 41 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "041" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 300

         'Suspensão = 50 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "050" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 400

         'Diferimento = 51 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "051" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 400

         'ICMS cobrado anteriormente por substituição tributária = 60 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "060" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 500  '500 – ICMS cobrado anteriormente por substituição tributária

         'Tributado com Permissão de Crédito e com Substituição tributária = 70 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "070" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 201

         'Outros  90 = CST_A
         If CODG_INTERNO_TRIBUTACAO = "090" Then _
            BUSCA_TRIBUTACAO_PRODUTO = 900
      'Case 2 Or 3 'CST x CSOSN (Cód. Do Regime normal x regime simples nacional)
      Case 3 'CST x CSOSN (Cód. Do Regime normal)

'estudar caso dessa rotina esta sendo feito na frmINTEGRA.PEDIDOitem_INTEGRA_MFI010

         '2 = Simples Nacional-Excesso de sublimite receita bruta
         '3 = Regime Normal - RPA

         BUSCA_TRIBUTACAO_PRODUTO = "" & CODG_INTERNO_TRIBUTACAO
   End Select

'esta fixo aqui porque ainda falta entendimento se a regra acima procede
'If BUSCA_TRIBUTACAO_PRODUTO <> 500 Then _
   BUSCA_TRIBUTACAO_PRODUTO = 400  'Não tributada pelo Simples Nacional

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "BUSCA_TRIBUTACAO_PRODUTO"
End Function

Public Function Checagem_Definição_Campo_Tabela(strBanco As String, _
                                                strTabela As String, _
                                                strCampo As String, _
                                                strTipoCampo As String, _
                                                strOperacao As Integer) As Boolean
'On Error GoTo ERRO_TRATA

   If Trim(strBanco) = "" Then _
      Exit Function

   If Trim(strTipoCampo) <> "" Then
      If EXISTE_OBJ_BANCO(strBanco, strTabela, "") = False Then
         If EXISTE_CAMPO_TABELA(strBanco, strCampo, strTabela) = False Then
            'alteração campo
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
   Checagem_Definição_Campo_Tabela = False
End Function

Public Sub Alteração_Definição_Campo_Tabela(strCampo As String, strTipoCampoAlter As String, strTabela As String, strBanco As String)
On Error Resume Next

   Dim TabTabela         As New ADODB.Recordset
   Dim strVACA As String

   If Trim(strCampo) = "" Then _
      Exit Sub

   If Trim(strTipoCampoAlter) = "" Then _
      Exit Sub

   CONT_N = 0

   If TabTabela.State = 1 Then _
      TabTabela.Close

   strVACA = ""

   strVACA = "select * from INFORMATION_SCHEMA.COLUMNS WITH (NOLOCK)"
   strVACA = strVACA & " WHERE COLUMN_NAME = '" & Trim(UCase(strCampo)) & "'"
   If Trim(strTabela) <> "" Then _
      strVACA = strVACA & " and table_NAME = '" & Trim(UCase(strTabela)) & "'"
   
'   TabTabela.Open strVACA, CONECTA_RETAGUARDA, , , adCmdText
   
   If Trim(strBanco) = "RETAGUARDA" Then _
      TabTabela.Open strVACA, CONECTA_RETAGUARDA, , , adCmdText

   If Trim(strBanco) = "SHFINFO" Then _
      TabTabela.Open strVACA, CONECTA_RETAGUARDA, , , adCmdText

   If Trim(strBanco) = "GLOBAL" Then _
      TabTabela.Open strVACA, CONECTA_GLOBAL, , , adCmdText
   
   While Not TabTabela.EOF
      If Trim(UCase(TabTabela.Fields("data_type").Value)) <> Trim(UCase(strTipoCampoAlter)) Then
         SQL = "ALTER TABLE " & Trim(TabTabela.Fields("table_name").Value) & " ALTER COLUMN " & strCampo & " " & Trim(strTipoCampoAlter)

         If Trim(strBanco) = "RETAGUARDA" Then _
            CONECTA_RETAGUARDA.Execute SQL

         If Trim(strBanco) = "SHFINFO" Then _
            CONECTA_RETAGUARDA.Execute SQL

         If Trim(strBanco) = "GLOBAL" Then _
            CONECTA_GLOBAL.Execute SQL

         CONT_N = CONT_N + 1
      End If

      TabTabela.MoveNext
      Err.Clear
   Wend
   If TabTabela.State = 1 Then _
      TabTabela.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "Alteração_Definição_Campo_Tabela"
End Sub

Public Function ESCOLHE_IMPRESSORA(Banco_Dados As String) As Boolean
On Error Resume Next

   ESCOLHE_IMPRESSORA = False

   If Trim(UCase(Banco_Dados)) = UCase("GLOBAL") Then
      Crystaldsn = Servidor_Global
      Crystaldsq = Banco_Dados
      Crystaluid = USUARIO_ADM_SQLSERVER
      Crystalpwd = SENHA_ADM_SQLSERVER
      Else
         Crystaldsn = SERVIDOR_MEGASIM
         Crystaldsq = Banco_Dados
         Crystaluid = USUARIO_ADM_SQLSERVER
         Crystalpwd = SENHA_ADM_SQLSERVER
   End If

   frmINICIO.Dialogo.CancelError = True
   frmINICIO.Dialogo.ShowPrinter

   If Err.Number > 0 Then
      ESCOLHE_IMPRESSORA = False
      'TRATA_ERROS Err.Description & "-" & Err.Number, "mdlGeral", "Alteração_Definição_Campo_Tabela"
      Else: ESCOLHE_IMPRESSORA = True
   End If

End Function

Sub TABELAS_RETAGUARDA()
'============TABELA PESSOA
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PESSOA", "") = False Then
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
   If EXISTE_OBJ_BANCO("RETAGUARDA", "EMPRESA", "") = False Then
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
   If EXISTE_OBJ_BANCO("RETAGUARDA", "USUARIO", "") = False Then
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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PRODUTO", "") = True Then _
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "STATUS", "PRODUTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN STATUS"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOITEM", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ", "PEDIDOITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM DROP COLUMN SEQ"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOITEM", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRECO_CUSTO", "PEDIDOITEM") = False Then
         SQL = "ALTER TABLE PEDIDOITEM ADD PRECO_CUSTO float"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
'==============================================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FONE", "") = True Then _
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "FONE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FONE DROP COLUMN EMPRESA_ID"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "BANCO", "") = True Then _
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "BANCO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE BANCO DROP COLUMN EMPRESA_ID"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "AGENCIA", "") = False Then
      SQL = "CREATE TABLE [dbo].[AGENCIA]("
      SQL = SQL & " [NUMR_AGENCIA] [varchar](10) NOT NULL,"
      SQL = SQL & " [BANCO] [int] NOT NULL,"
      SQL = SQL & " [NOME_AGENCIA] [varchar](100) NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
      Else
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "AGENCIA") = True Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE AGENCIA DROP COLUMN EMPRESA_ID"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CONTA", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "CONTA") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA ADD PESSOA_ID BIGINT"

         SQL = "update CONTA set pessoa_id = codg_cliente"
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA DROP COLUMN CODG_CLIENTE"
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "CONTA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CONTA DROP COLUMN EMPRESA_ID"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CHEQUE", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "CHEQUE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE DROP COLUMN EMPRESA_ID"
   
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_DOC", "CHEQUE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CHEQUE ADD NUMR_DOC NVARCHAR(30)"
   End If
'==============================================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOITEM", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "STATUS", "PEDIDOITEM") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE PEDIDOITEM ADD STATUS CHAR(1)"
      End If
   End If

'==================================================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CSOSN", "") = False Then
      SQL = " CREATE TABLE CSOSN("
      SQL = SQL & " CODIGO NVARCHAR(3) NOT NULL,"
      SQL = SQL & " DESCRICAO NVARCHAR(max) NOT NULL,"
      SQL = SQL & " OBS NVARCHAR(max) NULL,"
      SQL = SQL & " CONSTRAINT PK_TRIBUTACAO_CSOSN PRIMARY KEY CLUSTERED"
      SQL = SQL & " (Codigo Asc )"
      SQL = SQL & " WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 101
      SQL = SQL & ",'Tributada pelo Simples Nacional com permissão de crédito'"
      SQL = SQL & ",'classificam-se neste código as operações que permitem a indicação da alíquota de ICMS devido no Simples Nacional e o valor do crédito correspondente'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 102
      SQL = SQL & ",'Tributada pelo Simples Nacional sem permissão de crédito'"
      SQL = SQL & ",'classificam-se código as operações que não permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor do crédito, e não estejam abrangidas nas hipóteses dos códigos 103, 203, 300, 400, 500 e 900'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 103
      SQL = SQL & ",'Isenção do ICMS no Simples Nacional para faixa de receita bruta'"
      SQL = SQL & ",'classificam-se neste código as operações praticadas por optantes do Simples Nacional contempladas com isenção concedida para faixa de receita bruta nos termos da Lei Complementar n. 123 de 2006'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 201
      SQL = SQL & ",'Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações  que permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor crédito e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 202
      SQL = SQL & ",'Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações  que não permitem a indicação da alíquota do ICMS devido pelo Simples Nacional e do valor crédito, e não estejam abrangidas nas hipóteses dos códigos 103, 203, 300, 400, 500 e 900 e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 203
      SQL = SQL & ",'Isenção do ICMS no Simples Nacional para a faixa de receita bruta e com cobrança de ICMS por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações que praticadas por optantes do Simples Nacional contemplados com isenção para a faixa de receita bruta, mas com ICMS cobrado por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 300
      SQL = SQL & ",'Imune'"
      SQL = SQL & ",'classificam-se neste código as operações que praticadas por optantes do Simples Nacional contempladas com imunidade do ICMS'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 400
      SQL = SQL & ",'Não tributada pelo Simples Nacional'"
      SQL = SQL & ",'classificam-se neste código as operações que praticadas por optantes do Simples NacionaL não sujeitas à tributação pelo ICMS dentro do Simples Nacional'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 500
      SQL = SQL & ",'ICMS cobrado anteriormente por substituição tributária'"
      SQL = SQL & ",'classificam-se neste código as operações sujeitas exclusivamente ao regime de substituição tributária na condição de substituído tributário ou no caso de antecipações'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CSOSN ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 900
      SQL = SQL & ",'Outros'"
      SQL = SQL & ",'classificam-se neste código as operações que não se enquadrem nos códigos 101, 102, 103, 201, 202, 203, 300, 400 e 500'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CST", "") = False Then
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
      SQL = SQL & ",'Tributada  e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ",'Tributada  e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 20
      SQL = SQL & ",'Com redução de base de cálculo'"
      SQL = SQL & ",'Com redução de base de cálculo'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 30
      SQL = SQL & ",'Isenta ou não tributada e com cobrança do ICMS por substituição tributária'"
      SQL = SQL & ",'Isenta ou não tributada e com cobrança do ICMS por substituição tributária'"
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
      SQL = SQL & ",'Não tributada'"
      SQL = SQL & ",'Não tributada'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 50
      SQL = SQL & ",'Suspensão'"
      SQL = SQL & ",'Suspensão'"
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
      SQL = SQL & ",'ICMS cobrado anteriormente por substituição tributária'"
      SQL = SQL & ",'ICMS cobrado anteriormente por substituição tributária'"
      SQL = SQL & ")"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "insert into CST ("
         SQL = SQL & "codigo,descricao,obs"
      SQL = SQL & ")"
      SQL = SQL & " values("
      SQL = SQL & 70
      SQL = SQL & ",'Com redução de base de cálculo e cobrança de ICMS por substituição tributária'"
      SQL = SQL & ",'Com redução de base de cálculo e cobrança de ICMS por substituição tributária'"
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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ACESSO", "") = False Then
      SQL = "CREATE TABLE [dbo].[ACESSO]("
      SQL = SQL & " [USUARIO_ID] [int] NOT NULL,"
      SQL = SQL & " [PROGRAMA_ID] Not [Int]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

'=================================
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDO", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_CLIENTE", "PEDIDO") = True Then
         SQL = "ALTER TABLE PEDIDO ADD CLIENTE_ID BIGINT"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "update PEDIDO set "
         SQL = SQL & " cliente_id = codg_cliente"
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE PEDIDO DROP COLUMN CODG_CLIENTE"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   'INICIO EM 11/05/2012
   ' ESTE CAMPO FARA O CONTROLE DE FORMS DESENVOLVIDOS E SOMENTE UTILIZADOS POR AQUELA EMPRESA
   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CGCEMPRESA", "PROGRAMA") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PROGRAMA ADD CGCEMPRESA NVARCHAR(20) NULL"

   'FIM  11/05/2012
End Sub

Public Sub spPessoa(Acao_N As Long, ID_N As Long, CNPJCPF As String, DESCRICAO As String, RAZAO As String, SITUACAO As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   If PESSOA_ID_N <= 0 Then _
      PESSOA_ID_N = MAX_ID("PESSOA_ID", "PESSOA", "", "", "", "")

   DESCRICAO = Replace(DESCRICAO, "'", " ")
   RAZAO = Replace(RAZAO, "'", " ")
   CNPJCPF = Replace(CNPJCPF, "'", "")
   CNPJCPF = Replace(CNPJCPF, ".", "")
   CNPJCPF = Replace(CNPJCPF, "-", "")
   CNPJCPF = Replace(CNPJCPF, "/", "")

   SQL = "spPessoa " & Acao_N & "," & PESSOA_ID_N & ",'" & Trim(CNPJCPF) & "'" & ",'" & Trim(DESCRICAO) & "'" & ",'" & Trim(RAZAO) & "'" & ",'" & Trim(SITUACAO) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spPessoa"
End Sub

Public Sub spEmail(Acao_N As Long, ID_N As Long, EMAIL_A As String, PessoaID As Long)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   If PessoaID <= 0 Then
      MsgBox "Informar Pessoa_id."
      Exit Sub
   End If

   If Trim(EMAIL_A) = "" Then
      MsgBox "Informar email."
      Exit Sub
   End If

   If ID_N <= 0 Then _
      ID_N = MAX_ID("Email_ID", "Email", "", "", "", "")

   EMAIL_A = Replace(EMAIL_A, "'", " ")
   EMAIL_A = Replace(EMAIL_A, "'", "")
   'Email_A = Replace(Email_A, ".", "")
   'Email_A = Replace(Email_A, "-", "")
   EMAIL_A = Replace(EMAIL_A, "/", "")

   SQL = "spEmail " & Acao_N & "," & ID_N & ",'" & Trim(EMAIL_A) & "'" & ",'" & Trim(PessoaID) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spEmail"
End Sub

Public Sub spIE(Acao_N As Long, IE_ID_N As String, IE_A As String, END_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   'If IE_ID_N <= 0 Then _
      IE_ID_N = MAX_ID("IE_ID", "IE", "", "", "", "")

   IE_A = Replace(IE_A, "'", " ")
   IE_A = Replace(IE_A, "'", "")
   IE_A = Replace(IE_A, ".", "")
   IE_A = Replace(IE_A, "-", "")
   IE_A = Replace(IE_A, "/", "")

   SQL = "spIE " & Acao_N & "," & IE_ID_N & "," & PESSOA_ID_N & ",'" & Trim(IE_A) & "'," & END_ID_N

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spIE"
End Sub

Public Sub spIM(Acao_N As Long, IM_ID_N As Long, IM_A As String, END_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   If Trim(IM_A) = "" Then
      MsgBox "Informar IM."
      Exit Sub
   End If

   'If IM_ID_N <= 0 Then _
      IM_ID_N = MAX_ID("IM_ID", "IM", "", "", "", "")

   IM_A = Replace(IM_A, "'", " ")
   IM_A = Replace(IM_A, "'", "")
   IM_A = Replace(IM_A, ".", "")
   IM_A = Replace(IM_A, "-", "")
   IM_A = Replace(IM_A, "/", "")

   SQL = "spIM " & Acao_N & "," & IM_ID_N & "," & PESSOA_ID_N & ",'" & Trim(IM_A) & "'," & END_ID_N

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spIM"
End Sub

Public Sub spCliente(Acao_N As Long, TIPO_CLIENTE As String, CGCCPF As String, NOME As String, RAZAO_SOCIAL As String, _
                     DT_NASC As String, DT_CAD As String, STATUS As String, SEXO As String, CONTATO As String, _
                     REGIAO As String, ORIGEM As String, LIMITE_CREDITO As String, ESTRANGEIRO As String, _
                     CPFFUNCCONVENIO As String, PERC_DESC_CONVENIO As String, OBS As String, CODG_SUFRAMA As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedCLIENTEento."
      Exit Sub
   End If

   If Trim(DT_NASC) = "" Then
      DT_NASC = "01/01/1900 00:00:00"
   End If
   If Trim(REGIAO) = "" Then
      REGIAO = "NULL"
   End If

   SQL = "spCLIENTE " & Acao_N & "," & CLIENTE_ID_N & "," & PESSOA_ID_N & "," & ESTABELECIMENTO_ID_N & "," & VENDEDOR_ID_N & "," & _
                                     TIPO_CLIENTE & ",'" & CGCCPF & "','" & Trim(Replace(NOME, "'", " ")) & "','" & Trim(Replace(RAZAO_SOCIAL, "'", " ")) & "','" & DT_NASC & "','" & _
                                     DT_CAD & "','" & STATUS & "','" & SEXO & "','" & Trim(Replace(CONTATO, "'", " ")) & "'," & REGIAO & ",'" & _
                                     ORIGEM & "','" & tpMOEDA(LIMITE_CREDITO) & "'," & ESTRANGEIRO & ",'" & CPFFUNCCONVENIO & "','" & _
                                     tpMOEDA(PERC_DESC_CONVENIO) & "','" & Trim(Replace(OBS, "'", " ")) & "','" & Trim(Replace(CODG_SUFRAMA, "'", " ")) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spCLIENTE"
End Sub

Public Sub spPedidoComanda(Acao_N As Integer, Seq_Comanda_n As Long, Seq_Pedido_n As Long, ORIGEM_ITEM_A As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If
'set esta duplicando quando chama a comanda mais de uma vez no mesmo pedido
   If Acao_N <> 3 Then
      If PEDIDO_ID_N <= 0 Then
         MsgBox "Informar Pedido."
         Exit Sub
      End If
      If CARTAOBARRA_ID_N <= 0 Then
         MsgBox "Informar Comanda."
         Exit Sub
      End If

      Dim TabPedidoComanda As New ADODB.Recordset

      If TabPedidoComanda.State = 1 Then _
         TabPedidoComanda.Close
   
      SQL = "select * from PEDIDOCOMANDA WITH (NOLOCK)"
      SQL = SQL & " where cartaobarra_id = " & CARTAOBARRA_ID_N

If Trim(ORIGEM_ITEM_A) = "COMANDA" Then
   SQL = SQL & " and seq_COMANDA_id = " & Seq_Comanda_n
   Else: SQL = SQL & " and seq_PEDIDO_id = " & Seq_Pedido_n
End If

      TabPedidoComanda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabPedidoComanda.EOF Then
         Acao_N = 1
         Else: Acao_N = 2
      End If
      If TabPedidoComanda.State = 1 Then _
         TabPedidoComanda.Close
   End If   'If ACAO_N <> 3 Then

   SQL = "spPedidoComanda " & Acao_N & "," & PEDIDO_ID_N & "," & CARTAOBARRA_ID_N & "," & Seq_Comanda_n & "," & Seq_Pedido_n
   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spFONE"
End Sub

Public Sub spFONE(Acao_N As Long, ID_N As Long, NUMR_FONE As String, PessoaID As Long, DDD As Integer, LOCAL_A As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   If PessoaID <= 0 Then
      MsgBox "Informar Pessoa_id."
      Exit Sub
   End If

   If Trim(NUMR_FONE) = "" Then
      MsgBox "Informar FONE."
      Exit Sub
   End If

   If ID_N <= 0 Then _
      ID_N = MAX_ID("FONE_ID", "FONE", "", "", "", "")

   NUMR_FONE = Replace(NUMR_FONE, "'", " ")
   NUMR_FONE = Replace(NUMR_FONE, "'", "")
   'NUMR_FONE = Replace(NUMR_FONE, ".", "")
   'NUMR_FONE = Replace(NUMR_FONE, "-", "")
   NUMR_FONE = Replace(NUMR_FONE, "/", "")

   SQL = "spFONE " & Acao_N & "," & ID_N & ",'" & Trim(NUMR_FONE) & "'" & "," & Trim(PessoaID) & "," & Trim(DDD) & ",'" & Trim(LOCAL_A) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spFONE"
End Sub

Public Sub spCEP(Acao_N As Integer, CEP_ID As String, CIDADE As String, UF As String, IBGE_ID As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If
   If Trim(UF) = "" Then
      MsgBox "Informar UF."
      Exit Sub
   End If
   If Trim(CIDADE) = "" Then
      MsgBox "Informar CIDADE."
      Exit Sub
   End If
   If Trim(CEP_ID) = "" Then
      MsgBox "Informar CEP."
      Exit Sub
   End If
   If Trim(IBGE_ID) = "" Then
      IBGE_ID = 0
   End If

   CIDADE = Replace(CIDADE, "'", " ")
   CIDADE = Replace(CIDADE, "'", "")
   CIDADE = Replace(CIDADE, ".", "")
   CIDADE = Replace(CIDADE, "-", "")
   CIDADE = Replace(CIDADE, "/", "")

   SQL = "spCEP " & Acao_N & ",'" & CEP_ID & "','" & Trim(CIDADE) & "','" & Trim(UF) & "'," & Trim(IBGE_ID)

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spCEP"
End Sub

Public Sub spENDERECO(Acao_N As Integer, _
                      ENDERECO_ID_N As Long, _
                      PESSOAID_N As Long, _
                      CEP_ID_A As String, _
                      RUA_A As String, _
                      BAIRRO_A As String, _
                      COMPLEMENTO_A As String, _
                      TIPO_A As String, _
                      NUMERO_N As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If
   If Trim(CEP_ID_A) = "" Then
      MsgBox "Informar Cep."
      Exit Sub
   End If
   If Trim(PESSOAID_N) = "" Then
      MsgBox "Informar PESSOA."
      Exit Sub
   End If
   If Trim(CEP_ID_A) = "" Then
      MsgBox "Informar CEP."
      Exit Sub
   End If

   RUA_A = Replace(RUA_A, "'", " ")
   BAIRRO_A = Replace(BAIRRO_A, "'", " ")
   COMPLEMENTO_A = Replace(COMPLEMENTO_A, "'", " ")

   SQL = "spENDERECO " & Acao_N & _
             "," & ENDERECO_ID_N & _
             "," & PESSOAID_N & _
             ",'" & CEP_ID_A & "'" & _
             ",'" & RUA_A & "'" & _
             ",'" & BAIRRO_A & "'" & _
             ",'" & COMPLEMENTO_A & "'" & _
             ",'" & TIPO_A & "'" & _
             ",'" & NUMERO_N & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spENDERECO"
End Sub

Public Sub GRAVA_RG(NUMR_RG As String, ORGAO_A As String, DT_EXP_D As String)
'On Error GoTo ERRO_TRATA

   Dim TbRg As New ADODB.Recordset

   If Trim(DT_EXP_D) = "" Then _
      DT_EXP_D = 0

   If PESSOA_ID_N > 0 Then

      If TbRg.State = 1 Then _
         TbRg.Close

      SQL = "select * from RG WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TbRg.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TbRg.EOF Then
         SQL = "INSERT INTO RG "
            SQL = SQL & " (pessoa_id,numero_rg, orgao, dt_exp ) "
         SQL = SQL & " VALUES ("
            SQL = SQL & PESSOA_ID_N
            SQL = SQL & ",'" & Trim(NUMR_RG) & "'"
            SQL = SQL & ",'" & Trim(ORGAO_A) & "'"
            SQL = SQL & "," & DMA(DT_EXP_D)
         SQL = SQL & " )"
         Else
             SQL = "UPDATE RG SET "
             SQL = SQL & " numero_rg = '" & Trim(NUMR_RG) & "'"
             SQL = SQL & ", Orgao = '" & Trim(ORGAO_A) & "'"
             SQL = SQL & ", dt_exp = '" & DMA(DT_EXP_D) & "'"
             SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      End If
      If TbRg.State = 1 Then _
         TbRg.Close

      CONECTA_RETAGUARDA.Execute SQL

      'MsgBox "Processo de inclusão/alteração de RG realizado com sucesso."
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GRAVA_RG"
End Sub

Public Function GRAVA_FONE_PESSOA(NUMR_FONE_A As String, DDD_A As String, LOCAL_A As String) As Boolean
'On Error GoTo ERRO_TRATA

   Dim DDD_N As String

   GRAVA_FONE_PESSOA = False
   DDD_N = 0 & DDD_A

   If Trim(NUMR_FONE_A) <> "" And PESSOA_ID_N > 0 Then
      Dim TabFone  As New ADODB.Recordset

      If TabFone.State = 1 Then _
         TabFone.Close

      SQL = "select * from FONE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and numero = '" & Trim(NUMR_FONE_A) & "'"
      TabFone.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabFone.EOF Then
         spFONE 1, 0, Trim(NUMR_FONE_A), PESSOA_ID_N, Trim(DDD_N), Trim(LOCAL_A)
         Else: spFONE 2, TabFone.Fields("fone_id").Value, Trim(NUMR_FONE_A), PESSOA_ID_N, Trim(DDD_N), Trim(LOCAL_A)
      End If
      If TabFone.State = 1 Then _
         TabFone.Close

      GRAVA_FONE_PESSOA = True
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GRAVA_FONE_PESSOA"
End Function

Public Sub GRAVA_IE(NUMR_IE_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(NUMR_IE_A) <> "" And PESSOA_ID_N > 0 And ENDERECO_ID_N > 0 Then
      Dim tabIE As New ADODB.Recordset

      'NUMR_ID_N = 0
      'If IsNumeric(NUMR_IE_A) Then
      '   NUMR_ID_N = NUMR_IE_A
      '   If NUMR_ID_N = 0 Then _
      '      NUMR_IE_A = "ISENTO"
      'End If
      'NUMR_ID_N = 0

      NUMR_IE_A = "" & Trim(NUMR_IE_A)

      If tabIE.State = 1 Then _
         tabIE.Close

      SQL = "select * from IE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and ENDERECO_ID = " & ENDERECO_ID_N
      tabIE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If tabIE.EOF Then
         IE_ID = MAX_ID("IE_ID", "IE", "", "", "", "")

         SQL = "INSERT INTO IE "
            SQL = SQL & " (IE_ID, PESSOA_ID, Numr_IE,endereco_id) "
         SQL = SQL & " VALUES ("
            SQL = SQL & IE_ID
            SQL = SQL & "," & PESSOA_ID_N
            SQL = SQL & ",'" & Trim(NUMR_IE_A) & "'"
            SQL = SQL & "," & ENDERECO_ID_N
         SQL = SQL & ")"
         Else
            SQL = "UPDATE IE SET "
               SQL = SQL & " Numr_Ie = '" & Trim(NUMR_IE_A) & "'"
               SQL = SQL & ", ENDERECO_ID = " & ENDERECO_ID_N
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and ENDERECO_ID = " & ENDERECO_ID_N
      End If
      If tabIE.State = 1 Then _
         tabIE.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GRAVA_IE"
End Sub

Public Sub GRAVA_IM(NUMR_IM_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(NUMR_IM_A) <> "" And PESSOA_ID_N > 0 And ENDERECO_ID_N > 0 Then
      Dim tabIE As New ADODB.Recordset

      If tabIE.State = 1 Then _
         tabIE.Close

      SQL = "select * from IM WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and ENDERECO_ID = " & ENDERECO_ID_N
      tabIE.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If tabIE.EOF Then
         IM_ID = MAX_ID("IM_ID", "IM", "", "", "", "")

         SQL = "INSERT INTO IM "
            SQL = SQL & " (IM_ID, PESSOA_ID, Numr_IM,endereco_id) "
         SQL = SQL & " VALUES ("
            SQL = SQL & IM_ID
            SQL = SQL & "," & PESSOA_ID_N
            SQL = SQL & ",'" & Trim(NUMR_IM_A) & "'"
            SQL = SQL & "," & ENDERECO_ID_N
         SQL = SQL & ")"
         Else
            SQL = "UPDATE IE SET "
               SQL = SQL & " Numr_IM = '" & Trim(NUMR_IM_A) & "'"
               SQL = SQL & ", ENDERECO_ID = " & ENDERECO_ID_N
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and ENDERECO_ID = " & ENDERECO_ID_N
      End If
      If tabIE.State = 1 Then _
         tabIE.Close

      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GRAVA_IM"
End Sub

Public Sub ATUALIZA_TABELA_FAMILIAPRODUTO()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FAMILIAPRODUTO", "") = False Then
      SQL = "CREATE TABLE [dbo].[FAMILIAPRODUTO]("
      SQL = SQL & " [FAMILIAPRODUTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [CODG_FAMILIA] [nvarchar](10) NOT NULL,"
      SQL = SQL & " [DESCRICAO] [nvarchar](60) NOT NULL,"
      SQL = SQL & " [UNIDADE_MEDIDA] [nvarchar](4) NULL,"
      SQL = SQL & " [DESC_UNIDADE_MEDIDA] [nvarchar](50) NULL,"
      SQL = SQL & " [PERC_COMPOE_VENDA] [float]"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "FAMILIAPRODUTO_ID", "FAMILIAPRODUTO") = True Then
      Alteração_Definição_Campo_Tabela "FAMILIAPRODUTO_ID", "BIGINT NOT NULL", "FAMILIAPRODUTO", "RETAGUARDA"
      Else: CONECTA_RETAGUARDA.Execute "ALTER TABLE FAMILIAPRODUTO ADD FAMILIAPRODUTO_ID BIGINT NOT NULL"
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRODUCAO", "FAMILIAPRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE FAMILIAPRODUTO ADD PRODUCAO BIT"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_COMPOE_VENDA", "FAMILIAPRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE FAMILIAPRODUTO ADD PERC_COMPOE_VENDA FLOAT"

   SQL = "update FAMILIAPRODUTO set producao = 0"
   SQL = SQL & " where producao is null"
   CONECTA_RETAGUARDA.Execute SQL

   SQL = "update FAMILIAPRODUTO set perc_compoe_venda = 0"
   SQL = SQL & " where perc_compoe_venda  is null"
   CONECTA_RETAGUARDA.Execute SQL

   If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_FAMILIAPRODUTO", "") = False Then
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

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "nacional", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'PRODUTO.nacional'" & "," & "'ORIGEM_MERCADO'" & "," & "'COLUMN'"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_fornec", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'PRODUTO.codg_fornec'" & "," & "'FORNECEDOR_ID'" & "," & "'COLUMN'"
   
   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_PRODUTO", "PRODUTO") = True Then _
      Alteração_Definição_Campo_Tabela "CODG_PRODUTO", "NVARCHAR(100)", "PRODUTO", "RETAGUARDA"

   'If EXISTE_OBJ_BANCO("RETAGUARDA","pk_PRODUTO","") = True Then _
      CONECTA_RETAGUARDA.Execute SQL = "ALTER TABLE PRODUTO DROP CONSTRAINT pk_PRODUTO "

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRODUTO_ID", "PRODUTO") = True Then
      Alteração_Definição_Campo_Tabela "PRODUTO_ID", "BIGINT", "PRODUTO", "RETAGUARDA"
      Else: CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PRODUTO_ID BIGINT"
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESO_LIQUIDO", "PRODUTO") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PESO_LIQUIDO FLOAT"
      Else: Alteração_Definição_Campo_Tabela "PESO_LIQUIDO", "FLOAT", "PRODUTO", "RETAGUARDA"
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESO_BRUTO", "PRODUTO") = False Then
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PESO_BRUTO FLOAT"
      Else: Alteração_Definição_Campo_Tabela "PESO_BRUTO", "Float", "PRODUTO", "RETAGUARDA"
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "FAMILIAPRODUTO_ID", "PRODUTO") = True Then
      Alteração_Definição_Campo_Tabela "FAMILIAPRODUTO_ID", "BIGINT", "PRODUTO", "RETAGUARDA"
      Else: CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD FAMILIAPRODUTO_ID BIGINT"
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_PROD", "PRODUTO") = True Then _
      Alteração_Definição_Campo_Tabela "TIPO_PROD", "int", "PRODUTO", "RETAGUARDA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESCRICAO", "PRODUTO") = True Then _
      Alteração_Definição_Campo_Tabela "DESCRICAO", "nvarchar(250)", "PRODUTO", "RETAGUARDA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "REFERENCIA", "PRODUTO") = True Then _
      Alteração_Definição_Campo_Tabela "REFERENCIA", "NVARCHAR(MAX)", "PRODUTO", "RETAGUARDA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_BARRA", "PRODUTO") = True Then _
      Alteração_Definição_Campo_Tabela "CODG_BARRA", "NVARCHAR(MAX)", "PRODUTO", "RETAGUARDA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PATH_IMAGEM", "PRODUTO") = True Then _
      Alteração_Definição_Campo_Tabela "PATH_IMAGEM", "NVARCHAR(MAX)", "PRODUTO", "RETAGUARDA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "LOCACAO", "PRODUTO") = True Then _
      Alteração_Definição_Campo_Tabela "LOCACAO", "NVARCHAR(MAX)", "PRODUTO", "RETAGUARDA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_USUARIO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'PRODUTO.CODG_USUARIO'" & "," & "'USUARIO_ID'" & "," & "'COLUMN'"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "TAMANHO", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD TAMANHO BIGINT NULL"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "MARCA_ID", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD MARCA_ID BIGINT NULL"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRODUTO_BALANCA", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PRODUTO_BALANCA bit"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERMITE_DESCONTO", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PERMITE_DESCONTO BIT NULL"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CFOP", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN CFOP"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO_ORIGINAL", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN CODIGO_ORIGINAL"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALIQUOTA_SUBST_TRIBUTARIA", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN ALIQUOTA_SUBST_TRIBUTARIA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "VALOR_VENDA", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN VALOR_VENDA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "VALOR_CUSTO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN VALOR_CUSTO"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "GRUPO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN GRUPO"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "NATUREZA", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN NATUREZA"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CLASSIFICACAOFISCAL", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN CLASSIFICACAOFISCAL"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "IMAGEM", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN IMAGEM"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "STATUS", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN STATUS"
'=====================
   If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTDE", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN QTDE"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTDE_RETIDO", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN QTDE_RETIDO"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "qtd_alocada", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN qtd_alocada"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO_FABRICA", "PRODUTO") = True Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO DROP COLUMN CODIGO_FABRICA"
'===================
   If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONCEDER_PRODUCAO", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD CONCEDER_PRODUCAO BIT"

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_COMPOE_VENDA", "PRODUTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD PERC_COMPOE_VENDA float"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_PRODUTO", "") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE PRODUTO ADD CONSTRAINT pk_PRODUTO PRIMARY KEY (PRODUTO_ID)"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PRODUTO_EMPRESA", "") = False Then
      SQL = "ALTER TABLE [dbo].[PRODUTO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
      SQL = SQL & " References [dbo].[Empresa]([EMPRESA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTO] CHECK CONSTRAINT [FK_PRODUTO_EMPRESA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PRODUTO_FAMILIAPRODUTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[PRODUTO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTO_FAMILIAPRODUTO] FOREIGN KEY([FAMILIAPRODUTO_ID])"
      SQL = SQL & " References [dbo].[FAMILIAPRODUTO]([FAMILIAPRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[PRODUTO] CHECK CONSTRAINT [FK_PRODUTO_FAMILIAPRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   'SQL = "update PEDIDOITEM set tipo_reg = 'PC' "
   'SQL = SQL & " where tipo_reg Is Null "
   'CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "MDLGERAL", "CHECA_TABELA_PRODUTO"
End Sub

Public Sub CHECA_TABELA_ESTOQUE()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTOQUE", "") = False Then
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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ESTOQUE_ESTABELECIMENTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[ESTOQUE] "
      SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_ESTOQUE_ESTABELECIMENTO] "
      SQL = SQL & " FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[ESTOQUE] CHECK CONSTRAINT [FK_ESTOQUE_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ESTOQUE_PRODUTO", "") = False Then
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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_GERA_CODIGO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_GERA_CODIGO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_ADIANTA", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_INSERT_ADIANTA"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_CEP", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_INSERT_CEP"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_DISTRIBUICAO", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_INSERT_DISTRIBUICAO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_EQUIPE", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_INSERT_EQUIPE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_FONE", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_INSERT_FONE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_INSERT_PROD", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_INSERT_PROD"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_PROC_FONE", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_PROC_FONE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SP_UPDATE_CEP", "") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP PROCEDURE SP_UPDATE_CEP"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TipoEmpresaCliente", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TipoEmpresaCliente"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "EMPRESAACESSO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE EMPRESAACESSO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CSON", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE CSON"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ALIQUOTA", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ALIQUOTA"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ACESSO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ACESSO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ABCREL", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ABCREL"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "DEBITO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE DEBITO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ListaPrecos", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ListaPrecos"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "NATUREZAOPERACAOPRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE NATUREZAOPERACAOPRODUTO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CABCOMPR", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE CABCOMPR"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CABDEVENT", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE CABDEVENT"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CABDEVSAI", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE CABDEVSAI"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "BancoNaturezaOcorrencia", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE BancoNaturezaOcorrencia"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "BancoParametroRetorno", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE BancoParametroRetorno"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ALTERACAO_CLIENTE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ALTERACAO_CLIENTE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ALTERACAO_PRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ALTERACAO_PRODUTO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ITEMRECEBIMENTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ITEMRECEBIMENTO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ITENSCOM", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ITENSCOM"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "alteração_CLIENTE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE alteração_CLIENTE"

   'If EXISTE_OBJ_BANCO("RETAGUARDA", "ALIQUOTA_UF", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE ALIQUOTA_UF"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "BANCOFINA", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE BANCOFINA"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "alteração_PRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE alteração_PRODUTO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRel_POSICAOESTOQUE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE vwRel_POSICAOESTOQUE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "vwRel_EstoqueEntrada", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE vwRel_EstoqueEntrada"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "POSICAOESTOQUE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE POSICAOESTOQUE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "consulta_tributação", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view consulta_tributação"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CADASTRO_IVA_UF_PRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP table CADASTRO_IVA_UF_PRODUTO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "consulta familia", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view consulta familia"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "consulta_venda", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view consulta_venda"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "familia_consulta", "V") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP view familia_consulta"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "EMPRESAPARAMETRO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE EMPRESAPARAMETRO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "DEVOLUCAO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE DEVOLUCAO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "DEVOLUCAOTEMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE DEVOLUCAOTEMP"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TIPO60MESTRE", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60MESTRE"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TIPO60ANALITICO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60ANALITICO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TIPO60ITEM", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60ITEM"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TIPO60RESUMODIA", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60RESUMODIA"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TIPO60RESUMOMENSAL", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE TIPO60RESUMOMENSAL"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "NOTATEMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE NOTATEMP"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FUNCIONARIO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE FUNCIONARIO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FUNCIONARIOCONVENIO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE FUNCIONARIOCONVENIO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "GRUPOPRODUTO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE GRUPOPRODUTO"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "IMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE IMP"

   If EXISTE_OBJ_BANCO("RETAGUARDA", "SERIEIMP", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE SERIEIMP"
End Sub

Public Sub ATUALIZA_TABELA_PESSOA()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PESSOA", "") = False Then
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
         If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_PESSOA", "") = False Then
            SQL = "ALTER TABLE PESSOA ADD CONSTRAINT pk_PESSOA PRIMARY KEY (PESSOA_ID)"
            CONECTA_RETAGUARDA.Execute SQL
         End If
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_REG", "PESSOA") = True Then _
            CONECTA_RETAGUARDA.Execute "ALTER TABLE PESSOA DROP COLUMN TIPO_REG"
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "PESSOATIPO", "") = False Then
   End If
   'CREATE INDEX IX_CNPJCPF ON PESSOA(CNPJCPF) WITH (ONLINE=ON, SORT_IN_TEMPDB=ON)
End Sub

Public Sub ATUALIZA_TABELA_EMPRESA()

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTABELECIMENTOACESSO", "") = False Then
      SQL = "CREATE TABLE [dbo].[ESTABELECIMENTOACESSO]("
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] Not [Int]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "EMPRESA", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Empresa_Optante_Simples", "EMPRESA") = True Then
         Alteração_Definição_Campo_Tabela "Empresa_Optante_Simples", "INT", "EMPRESA", "RETAGUARDA"
         SQL = "EXEC sp_rename " & "'EMPRESA.Empresa_Optante_Simples'" & "," & "'CRT'" & "," & "'COLUMN'"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "update empresa set crt = 1 where empresa_id = " & EMPRESA_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PAR", "EMPRESA") = True Then _
         Alteração_Definição_Campo_Tabela "PAR", "NVARCHAR(MAX)", "EMPRESA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INDR_INDUSTRIA", "EMPRESA") = True Then _
         Alteração_Definição_Campo_Tabela "INDR_INDUSTRIA", "BIT", "EMPRESA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "EMPRESA") = True Then _
         Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT NOT NULL", "EMPRESA", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codigo_cliente", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN codigo_cliente"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "caminhonfe", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN caminhonfe"
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "caminhonfe", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD caminhonfe NVARCHAR(50)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FONE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN FONE"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_CONSULTA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_CONSULTA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_FUNCIONARIO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_FUNCIONARIO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ECF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ECF"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "NUMR_PEDIDO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN NUMR_PEDIDO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_NOTA_ENTRADA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_NOTA_ENTRADA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_PEDIDO_ENTRADA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_PEDIDO_ENTRADA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTROLA_ESTOQUE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN CONTROLA_ESTOQUE"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALIQUOTA_DENTRO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ALIQUOTA_DENTRO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALIQUOTA_FORA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ALIQUOTA_FORA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_ICMS_ENTRADA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PERC_ICMS_ENTRADA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SENHA_BANCO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SENHA_BANCO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USUARIO_BANCO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN USUARIO_BANCO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGOBARRAPESOQTD", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN CODIGOBARRAPESOQTD"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CAMINHOREL", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN CAMINHOREL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_DESCONTO_CONV", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PERC_DESCONTO_CONV"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PSPermiteAlterarValor", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSPermiteAlterarValor"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPOssinatura", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN TIPOssinatura"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VersaoPaf", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN VersaoPaf"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PCNumeroECF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCNumeroECF"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "desabilitareducaoz", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN desabilitareducaoz"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PCUsaBalancaCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCUsaBalancaCaixa"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PSAntesIniciarVendaCartao", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAntesIniciarVendaCartao"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PCUsaLeitorSerial", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCUsaLeitorSerial"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PCImpressoraUsaPorta", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCImpressoraUsaPorta"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PCPortaBalancaCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCPortaBalancaCaixa"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PCPorta", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCPorta"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PCPortaLeitorSerial", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PCPortaLeitorSerial"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PSAliquotas", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAliquotas"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PBDBancoDadosCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PBDBancoDadosCaixa"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PSSistemaControlaCaixa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSSistemaControlaCaixa"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PSAntesIniciarVendaCPF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAntesIniciarVendaCPF"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PTBandeiraTecBan", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraTecBan"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PTBandeiraVisa", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraVisa"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PTBandeiraMasterCard", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraMasterCard"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PTBandeiraAmericanExpress", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTBandeiraAmericanExpress"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PSPermiteDesconto", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSPermiteDesconto"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PSAbrirGaveta", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PSAbrirGaveta"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PTUsaTef", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PTUsaTef"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ImprimeOrcNormalFiscal", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ImprimeOrcNormalFiscal"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VERSAO_CNIECF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN VERSAO_CNIECF"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Aliquota", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN Aliquota"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TipoTefDiscadoDedicado", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN TipoTefDiscadoDedicado"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Empresa_Regime_TARE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN Empresa_Regime_TARE"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "super_simples", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN super_simples"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "OPTANTE_TARE", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN OPTANTE_TARE"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PAR", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD PAR NVARCHAR(8) "

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "PIS", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD PIS float"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "COFINS", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD COFINS float"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CST_PIS", "empresa") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE empresa DROP COLUMN CST_PIS"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CST_COFINS", "empresa") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE empresa DROP COLUMN CST_COFINS"
'============================================================ DROP MIGRA ESTABELECIMENTO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "IMPRESSORA_FISCAL", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN IMPRESSORA_FISCAL"
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "IMPRESSORA_FISCAL", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD IMPRESSORA_FISCAL NVARCHAR(8)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "usa_impfiscal", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN usa_impfiscal"
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "usa_impfiscal", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD usa_impfiscal NVARCHAR(8)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "indr_industria", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN indr_industria"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LIBERA_DESCONTO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN LIBERA_DESCONTO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "usa_nfe", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN usa_nfe"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "RECEBE_PEDIDO_VENDA", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN RECEBE_PEDIDO_VENDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTOQUE_NEGATIVO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ESTOQUE_NEGATIVO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USA_TEF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN USA_TEF"
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "USA_TEF", "EMPRESA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD USA_TEF NVARCHAR(8) "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LEI_12741", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN LEI_12741"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "UsaCobBancaria", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN UsaCobBancaria"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INSTRUCAO_BOLETO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN INSTRUCAO_BOLETO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INSTRUCAO_FISCO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN INSTRUCAO_FISCO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ATUALIZA_ESTOQUE_REQ", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN ATUALIZA_ESTOQUE_REQ"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_JUROS_ATRAZO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN PERC_JUROS_ATRAZO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTD_DIAS_ATRAZO", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN QTD_DIAS_ATRAZO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "le_deposito", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN LE_DEPOSITO"
'============================================================
      CONT_N = 0
      If EMPRESA_ID_N <= 0 Then _
         EMPRESA_ID_N = 1
      If ESTABELECIMENTO_ID_N <= 0 Then _
         ESTABELECIMENTO_ID_N = 1

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "cgc", "EMPRESA") = True Then
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_REQORC", "EMPRESA") = True Then
         SQL = "EXEC sp_rename " & "'EMPRESA.SEQ_REQORC'" & "," & "'SEQ_PEDIDO'" & "," & "'COLUMN'"
         CONECTA_RETAGUARDA.Execute SQL

         Alteração_Definição_Campo_Tabela "SEQ_PEDIDO", "BIGINT", "EMPRESA", "RETAGUARDA"
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_EMPRESA", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA ADD CONSTRAINT pk_EMPRESA PRIMARY KEY (EMPRESA_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_EMPRESA_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[EMPRESA]  WITH CHECK ADD  CONSTRAINT [FK_EMPRESA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[EMPRESA] CHECK CONSTRAINT [FK_EMPRESA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
End Sub

Public Sub ATUALIZA_TABELA_USUARIO()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "USUARIO", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "USUARIO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO ADD PESSOA_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT", "USUARIO", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USUARIO_ID", "USUARIO") = True Then _
         Alteração_Definição_Campo_Tabela "USUARIO_ID", "BIGINT NOT NULL", "USUARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_DESCONTO", "USUARIO") = True Then _
         Alteração_Definição_Campo_Tabela "PERC_DESCONTO", "FLOAT", "USUARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_COMISSAO", "USUARIO") = True Then _
         Alteração_Definição_Campo_Tabela "PERC_COMISSAO", "FLOAT", "USUARIO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FUNCIONARIO", "USUARIO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO ADD FUNCIONARIO BIT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "nivel", "USUARIO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO DROP column nivel"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "empresa", "USUARIO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO DROP column empresa"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_USUARIO", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE USUARIO ADD CONSTRAINT pk_USUARIO PRIMARY KEY (USUARIO_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_USUARIO_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[USUARIO]  WITH CHECK ADD  CONSTRAINT [FK_USUARIO_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[USUARIO] CHECK CONSTRAINT [FK_USUARIO_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If
End Sub

Public Sub ATUALIZA_ESTABELECIMENTO()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "MENSAGEMSEFAZ", "") = False Then
      SQL = "CREATE TABLE [dbo].[MENSAGEMSEFAZ]("
      SQL = SQL & " [ERRO_ID] [bigint] NULL,"
      SQL = SQL & " [TIPO] [nchar](10) NULL,"
      SQL = SQL & " [MOTIVO] [nvarchar](max) NULL"
      SQL = SQL & " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ESTABELECIMENTO", "") = True Then

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_TEF", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO DROP COLUMN TIPO_TEF "
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "TIPO_TEF", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD TIPO_TEF NVARCHAR(1)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USA_TEF", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN USA_TEF"
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "USA_TEF", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD USA_TEF NVARCHAR(1) "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "IMPRESSORA_FISCAL", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN IMPRESSORA_FISCAL"
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "IMPRESSORA_FISCAL", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD IMPRESSORA_FISCAL NVARCHAR(1)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "usa_impfiscal", "EMPRESA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN usa_impfiscal"
      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "usa_impfiscal", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD usa_impfiscal NVARCHAR(1)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CNPJCPF", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CNPJCPF nvarchar(14)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DescRevenda", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO DROP column DescRevenda "

      If EXISTE_OBJ_BANCO("RETAGUARDA", "IX_ESTABELECIMENTO", "") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO DROP CONSTRAINT IX_ESTABELECIMENTO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INTEGRA", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO drop column INTEGRA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_FOLHA", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO DROP COLUMN CODG_FOLHA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "servidor", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO drop column servidor"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "nome_banco", "ESTABELECIMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO drop column nome_banco"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTABELECIMENTO_ID", "ESTABELECIMENTO") = True Then _
         Alteração_Definição_Campo_Tabela "ESTABELECIMENTO_ID", "BIGINT NOT NULL", "ESTABELECIMENTO", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LOCALIZACAO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LOCALIZACAO NVARCHAR(MAX)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTROLE_ESTOQUE", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CONTROLE_ESTOQUE BIT"

'==========================================================================================================
'INDR_INDUSTRIA
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INDR_INDUSTRIA", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD INDR_INDUSTRIA bit "

'LIBERA_DESCONTO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LIBERA_DESCONTO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LIBERA_DESCONTO bit "

'USA_NFe
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "USA_NFe", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD USA_NFe bit "

'RECEBE_PEDIDO_VENDA
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "RECEBE_PEDIDO_VENDA", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD RECEBE_PEDIDO_VENDA bit "

'ESTOQUE_NEGATIVO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ESTOQUE_NEGATIVO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD ESTOQUE_NEGATIVO bit "

'LEI_12741
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LEI_12741", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LEI_12741 bit "

'INSTRUCAO_BOLETO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INSTRUCAO_BOLETO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD INSTRUCAO_BOLETO NVARCHAR(MAX)"

'ATUALIZA_ESTOQUE_REQ
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ATUALIZA_ESTOQUE_REQ", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD ATUALIZA_ESTOQUE_REQ BIT "

'PERC_JUROS_ATRAZO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_JUROS_ATRAZO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD PERC_JUROS_ATRAZO FLOAT "

'QTD_DIAS_ATRAZO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "QTD_DIAS_ATRAZO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD QTD_DIAS_ATRAZO INT"

'VLR_DIA_COMPRA_PROD
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VLR_DIA_COMPRA_PROD", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD VLR_DIA_COMPRA_PROD FLOAT"

'chkPanific = MULT_EMPRESA_B
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "INDR_PANIFIC", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD INDR_PANIFIC bit"

'txtCasaInicioCodgProdBarra = CasaInicioCodgProdBarra
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CasaInicioCodgProdBarra", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CasaInicioCodgProdBarra int"

'txtTamanhoCodgProdBarra = TamanhoCodgProdBarra
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TamanhoCodgProdBarra", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD TamanhoCodgProdBarra int"

'txtTamanhoPesoValorBarra = TamanhoPesoValorBarra
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "TamanhoPesoValorBarra", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD TamanhoPesoValorBarra int"

'optgramas valor = peso_valor
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESO_VALOR", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD PESO_VALOR nvarchar(10) "

'AT_VENDA_MKP
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "AT_VENDA_MKP", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD AT_VENDA_MKP bit"

'DESCONTO_CLIENTE
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESCONTO_CLIENTE", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD DESCONTO_CLIENTE bit"

'DESCONTO_FUNCIONARIO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DESCONTO_FUNCIONARIO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD DESCONTO_FUNCIONARIO bit"

'chkPercDesconto
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LiberaPercDesconto", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LiberaPercDesconto bit"

'txtCodgProdutoReserva = CODG_PROD_RESERVA
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_PROD_RESERVA", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CODG_PROD_RESERVA INT "

'txtDiasAtrazo = DiasAtrazoCliente
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DiasAtrazoCliente", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD DiasAtrazoCliente INT "

'LIMPA_PEDIDO
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "LIMPA_PEDIDO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD LIMPA_PEDIDO bit "

'SEQ_CUPOM
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "SEQ_CUPOM", "ESTABELECIMENTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD SEQ_CUPOM BIGINT"
         SQL = "update ESTABELECIMENTO set seq_cupom = 0 "
         CONECTA_RETAGUARDA.Execute SQL
      End If

'seq_nota_saida
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "seq_nota_saida", "EMPRESA") = True Then
         'CRIANDO O CAMPO NA TABELA ESTABELECIMENTO
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "seq_nota_saida", "ESTABELECIMENTO") = False Then
            CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD SEQ_NOTA_SAIDA BIGINT"

            SQL = "update ESTABELECIMENTO set ESTABELECIMENTO.SEQ_NOTA_SAIDA = EMPRESA.SEQ_NOTA_SAIDA"
            SQL = SQL & " from ESTABELECIMENTO "
            SQL = SQL & " INNER JOIN EMPRESA"
            SQL = SQL & " ON ESTABELECIMENTO.EMPRESA_ID = EMPRESA.EMPRESA_ID"
            SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            CONECTA_RETAGUARDA.Execute SQL
         End If

         CONECTA_RETAGUARDA.Execute "ALTER TABLE EMPRESA DROP COLUMN SEQ_NOTA_SAIDA"
      End If
'==========================================================================================================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VERSAO_APLICATIVO", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD VERSAO_APLICATIVO nvarchar(20)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CARTAOADM_ID", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CARTAOADM_ID INT"
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ESTAB_CARTAOADM", "") = False Then
         SQL = "ALTER TABLE [dbo].[ESTABELECIMENTO]  WITH CHECK ADD  CONSTRAINT [FK_ESTAB_CARTAOADM] FOREIGN KEY([CARTAOADM_ID])"
         SQL = SQL & " References [dbo].[CARTAOADM]([CARTAOADM_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[ESTABELECIMENTO] CHECK CONSTRAINT [FK_ESTAB_CARTAOADM]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CSC", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD CSC nvarchar(30)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE ESTABELECIMENTO ADD CONSTRAINT pk_ESTABELECIMENTO PRIMARY KEY (ESTABELECIMENTO_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ESTABELECIMENTO_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[ESTABELECIMENTO]  WITH CHECK ADD  CONSTRAINT [FK_ESTABELECIMENTO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[Empresa]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[ESTABELECIMENTO] CHECK CONSTRAINT [FK_ESTABELECIMENTO_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

'DOC_FISCAL
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DOC_FISCAL", "ESTABELECIMENTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD DOC_FISCAL bit "
   
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ALTERA_FATURA", "ESTABELECIMENTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD ALTERA_FATURA bit "
         SQL = "update ESTABELECIMENTO set altera_fatura = 1 "
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If EXISTE_CAMPO_TABELA("RETAGUARDA", "USA_TAB_PRECO", "ESTABELECIMENTO") = False Then _
      CONECTA_RETAGUARDA.Execute "ALTER TABLE ESTABELECIMENTO ADD USA_TAB_PRECO bit "

   If EXISTE_OBJ_BANCO("RETAGUARDA", "EVENTO", "") = False Then
      SQL = "CREATE TABLE [dbo].[EVENTO]("
      SQL = SQL & " [EVENTO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ESTABELECIMENTO_ID] [int] NOT NULL,"
      SQL = SQL & " [USUARIO_ID] [BIGint] NOT NULL,"
      SQL = SQL & " [DT_EVENTO] [datetime] NOT NULL,"
      SQL = SQL & " [HISTORICO] [nvarchar](Max)"
      SQL = SQL & " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[EVENTO]  WITH CHECK ADD  CONSTRAINT [FK_EVENTO_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
      SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[EVENTO] CHECK CONSTRAINT [FK_EVENTO_ESTABELECIMENTO]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[EVENTO]  WITH CHECK ADD  CONSTRAINT [FK_EVENTO_USUARIO] FOREIGN KEY([USUARIO_ID])"
      SQL = SQL & " References [dbo].[USUARIO]([USUARIO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[EVENTO] CHECK CONSTRAINT [FK_EVENTO_USUARIO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
End Sub

Public Sub ATUALIZA_TABELA_ENTREGA()
   If EXISTE_OBJ_BANCO("RETAGUARDA", "ENTREGA", "") = False Then
      SQL = "CREATE TABLE [dbo].[ENTREGA]("
      SQL = SQL & " [ENTREGA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PESSOA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [EMPRESA_ID] [int] NOT NULL,"
      SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
      SQL = SQL & " [DT_AGENDA] [datetime] NULL,"
      SQL = SQL & " [DT_ENTREGA] [datetime] NULL,"
      SQL = SQL & " [CEP_ID] [NVARCHAR](8) NULL,"
      SQL = SQL & " [RUA] [NVARCHAR](50) NULL,"
      SQL = SQL & " [COMPLEMENTO] [NVARCHAR](50) NULL,"
      SQL = SQL & " [BAIRRO] [NVARCHAR](50) NULL,"
      SQL = SQL & " [CIDADE] [NVARCHAR](50) NULL,"
      SQL = SQL & " [UF] [NVARCHAR](2) NULL,"
      SQL = SQL & " [ATENDENTE_ID] [bigint] ,"
      SQL = SQL & " [ENTREGADOR_ID] [bigint] ,"

      SQL = SQL & " [ATENDENTE] [NVARCHAR](30) ,"
      SQL = SQL & " [ENTREGADOR] [NVARCHAR](30) ,"
      
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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "ENTREGA", "") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CEP", "ENTREGA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'ENTREGA.CEP'" & "," & "'CEP_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENTREGA_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ENTREGA_ID BIGINT NOT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD PESSOA_ID BIGINT NOT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PEDIDO_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD PEDIDO_ID BIGINT NOT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_CAD", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD DT_CAD DATETIME NOT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_AGENDA", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD DT_AGENDA DATETIME "
   
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DT_ENTREGA", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD DT_ENTREGA DATETIME "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CEP_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD CEP_ID NVARCHAR(8)"
   
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "RUA", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD RUA NVARCHAR(50)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "COMPLEMENTO", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD COMPLEMENTO NVARCHAR(50)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "BAIRRO", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD BAIRRO NVARCHAR(50)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CIDADE", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD CIDADE NVARCHAR(50)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "UF", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD UF NVARCHAR(2)"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENTREGADOR_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ENTREGADOR_ID BIGINT "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ENTREGADOR", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ENTREGADOR NVARCHAR(30) "

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "MONTADOR_ID", "ENTREGA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA DROP COLUMN MONTADOR_ID"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "MONTADOR", "ENTREGA") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA DROP COLUMN MONTADOR"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ATENDENTE_ID", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ATENDENTE_ID BIGINT "
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "ATENDENTE", "ENTREGA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE ENTREGA ADD ATENDENTE NVARCHAR(30) "

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_ENTREGA", "") = False Then
         SQL = "ALTER TABLE ENTREGA ADD CONSTRAINT pk_ENTREGA PRIMARY KEY (ENTREGA_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ENTREGA_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ENTREGA_PEDIDO", "") = False Then
         SQL = "ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
         SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_PEDIDO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_ENTREGA_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[ENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_ENTREGA_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = "ALTER TABLE [dbo].[ENTREGA] CHECK CONSTRAINT [FK_ENTREGA_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "OBSENTREGA", "") = False Then
      SQL = "CREATE TABLE [dbo].[OBSENTREGA]("
      SQL = SQL & " [OBSENTREGA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [ENTREGA_ID] [bigint] NOT NULL,"
      SQL = SQL & " [OBS] [nvarchar](max) NOT NULL,"
      SQL = SQL & " [DT_CAD] [datetime] NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_OBSENTREGA] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([OBSENTREGA_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) "
      SQL = SQL & " ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OBSENTREGA]  WITH CHECK ADD  CONSTRAINT [FK_OBSENTREGA_ENTREGA] FOREIGN KEY([ENTREGA_ID])"
      SQL = SQL & " References [dbo].[ENTREGA]([ENTREGA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[OBSENTREGA] CHECK CONSTRAINT [FK_OBSENTREGA_ENTREGA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
End Sub

Sub GRAVA_PRIMEIRA_DATA(CRYPTO_A As String)
'On Error Resume Next

   If Trim(CRYPTO_A) = "" Then _
      Exit Sub

   SQL = "update EMPRESA set "

   SQL = SQL & " par = '" & Trim(CRYPTO_A) & "'"

   SQL = SQL & " from EMPRESA "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL
End Sub

Public Sub GRAVA_PESSOA_ENDERECO()
'On Error GoTo ERRO_TRATA
'fazer normalizaão banco
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

   SQL = "select * from PESSOAENDERECO WITH (NOLOCK)"
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
      ENDERECO_ID_N = TabPE!ENDERECO_ID

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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "COMISSAOITEM", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE COMISSAOITEM"
   If EXISTE_OBJ_BANCO("RETAGUARDA", "COMISSAO", "U") = True Then _
      CONECTA_RETAGUARDA.Execute "DROP TABLE COMISSAO"
   
   If EXISTE_OBJ_BANCO("RETAGUARDA", "COMISSAO", "U") = False Then
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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_COMISSAO_EMPRESA", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAO]  WITH CHECK ADD  CONSTRAINT [FK_COMISSAO_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
      SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "ALTER TABLE [dbo].[COMISSAO] CHECK CONSTRAINT [FK_COMISSAO_EMPRESA]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "COMISSAOITEM", "U") = False Then
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

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_COMISSAOITEM_COMISSAO", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD CONSTRAINT [FK_COMISSAOITEM_COMISSAO] FOREIGN KEY([COMISSAO_ID])"
      SQL = SQL & " References [dbo].[COMISSAO]([COMISSAO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_COMISSAO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_COMISSAOITEM_PEDIDO", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD CONSTRAINT [FK_COMISSAOITEM_PEDIDO] FOREIGN KEY([PEDIDO_ID])"
      SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_PEDIDO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_COMISSAOITEM_VENDEDOR", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_COMISSAOITEM_VENDEDOR] FOREIGN KEY([VENDEDOR_ID])"
      SQL = SQL & " References [dbo].[VENDEDOR]([VENDEDOR_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_VENDEDOR]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_COMISSAOITEM_PRODUTO", "") = False Then
      SQL = "ALTER TABLE [dbo].[COMISSAOITEM]  WITH CHECK ADD  CONSTRAINT [FK_COMISSAOITEM_PRODUTO] FOREIGN KEY([PRODUTO_ID])"
      SQL = SQL & " References [dbo].[PRODUTO]([PRODUTO_ID])"
      CONECTA_RETAGUARDA.Execute SQL

      SQL = " ALTER TABLE [dbo].[COMISSAOITEM] CHECK CONSTRAINT [FK_COMISSAOITEM_PRODUTO]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_COMISSAOITEM_CLIENTE", "") = False Then
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
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CARTAOADM", "U") = False Then
      SQL = "CREATE TABLE [dbo].[CARTAOADM]("
      SQL = SQL & " [CARTAOADM_ID] [int] NOT NULL,"
      SQL = SQL & " [RAZAO] [nvarchar](70) NOT NULL,"
      SQL = SQL & " [FANTASIA] [nvarchar](70) NOT NULL,"
      SQL = SQL & " [CNPJ] [nvarchar](14) NOT NULL,"
      SQL = SQL & " [STATUS] [nchar](1) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_CARTAOADM] PRIMARY KEY CLUSTERED([CARTAOADM_ID] Asc)"
      SQL = SQL & " WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "CARTAOPEDIDO", "U") = False Then
      SQL = "CREATE TABLE [dbo].[CARTAOPEDIDO]("
      SQL = SQL & " [CARTAOPEDIDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [PEDIDO_ID] [bigint] NOT NULL,"
      SQL = SQL & " [BANDEIRA_ID] [char](2) NULL,"
      SQL = SQL & " [CNPJ_CARTAO] [nvarchar](14) NULL,"
      SQL = SQL & " [NUMR_AUTORIZACAO] [nchar](30) NULL"
      SQL = SQL & " ) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CARTAOPEDIDO_PEDIDO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CARTAOPEDIDO] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_CARTAOPEDIDO_PEDIDO] "
         SQL = SQL & " FOREIGN KEY([PEDIDO_ID])"
         SQL = SQL & " References [dbo].[PEDIDO]([PEDIDO_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[CARTAOPEDIDO] CHECK CONSTRAINT [FK_CARTAOPEDIDO_PEDIDO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "FORMAPAGTO", "U") = True Then

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "Cadastra_Impressora", "FORMAPAGTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORMAPAGTO DROP COLUMN Cadastra_Impressora"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORMA_ID", "FORMAPAGTO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'FORMAPAGTO.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORMA_ID", "ITEMLANCAMENTO") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'ITEMLANCAMENTO.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTAB_TESORA", "FORMAPAGTO") = False Then _
         If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTABILIZA", "FORMAPAGTO") = True Then _
            CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'FORMAPAGTO.CONTABILIZA'" & "," & "'CONTAB_TESORA'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTABILIZA", "FORMAPAGTO") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORMAPAGTO DROP COLUMN CONTABILIZA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTAB_BALCAO", "FORMAPAGTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORMAPAGTO ADD CONTAB_BALCAO BIT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "BAIXAAUTO", "FORMAPAGTO") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORMAPAGTO ADD BAIXAAUTO BIT NULL"

         SQL = "update FORMAPAGTO set"
         SQL = SQL & " baixaauto = 0 "
         SQL = SQL & " where baixaauto is null"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FUNC", "FORMAPAGTO") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE FORMAPAGTO ADD FUNC BIT NULL"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_FORMAPAGTO", "") = False Then
         SQL = "ALTER TABLE FORMAPAGTO ADD CONSTRAINT pk_FORMAPAGTO PRIMARY KEY (FORMAPAGTO_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_FORMAPAGTO_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[FORMAPAGTO] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_FORMAPAGTO_EMPRESA] "
         SQL = SQL & " FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[FORMAPAGTO] CHECK CONSTRAINT [FK_FORMAPAGTO_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORMA_ID", "caixatesorariaitem") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'caixatesorariaitem.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORMA_ID", "CAIXADIAITEM") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CAIXADIAITEM.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "TIPOVENDA", "U") = True Then
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "FORMA_ID", "TIPOVENDA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'TIPOVENDA.FORMA_ID'" & "," & "'FORMAPAGTO_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CONTABILIZA", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD CONTABILIZA BIT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_CARTAO_DEBITO", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD PERC_CARTAO_DEBITO FLOAT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERC_CARTAO_CREDITO", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD PERC_CARTAO_CREDITO FLOAT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PERMITE_DESCONTO", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD PERMITE_DESCONTO BIT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CC_ID", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD CC_ID INT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PreFatura", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD PreFatura BIT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PAGAR", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD PAGAR BIT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "RECEBER", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD RECEBER BIT NULL"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CARTAOADM_ID", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD CARTAOADM_ID INT"

      'If EXISTE_CAMPO_TABELA("RETAGUARDA", "PRAZO", "TIPOVENDA") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'TIPOVENDA.PRAZO'" & "," & "'DIASPRAZO'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "DIAVENCTO", "TIPOVENDA") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE TIPOVENDA ADD DIAVENCTO INT"

      SQL = "update TIPOVENDA set "
      SQL = SQL & " contabiliza = 1"
      SQL = SQL & " where contabiliza is null "
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "update TIPOVENDA set "
      SQL = SQL & " PermiteParcelar = 1"
      SQL = SQL & " where PermiteParcelar is null "
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "update TIPOVENDA set "
      SQL = SQL & " RECEBER = 1"
      SQL = SQL & " where RECEBER is null "
      CONECTA_RETAGUARDA.Execute SQL

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_TIPOVENDA", "") = False Then
         SQL = "ALTER TABLE TIPOVENDA ADD CONSTRAINT pk_TIPOVENDA PRIMARY KEY (TIPOVENDA_ID)"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_TIPOVENDA_EMPRESA", "") = False Then
         SQL = "ALTER TABLE [dbo].[TIPOVENDA] "
         SQL = SQL & " WITH CHECK ADD  CONSTRAINT [FK_TIPOVENDA_EMPRESA] "
         SQL = SQL & " FOREIGN KEY([EMPRESA_ID])"
         SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         CONECTA_RETAGUARDA.Execute SQL

         SQL = " ALTER TABLE [dbo].[TIPOVENDA] CHECK CONSTRAINT [FK_TIPOVENDA_EMPRESA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_TIPOVENDA_FORMAPAGTO", "") = False Then
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
         SQL = SQL & " [DIAVENCTO] [int] NULL,"
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

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then _
      Exit Sub

   If TabTemp.State = 1 Then _
      TabTemp.Close

   SQL = "select * from ObservacaoNota "
   SQL = SQL & " where codigo = '900'"
   TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If TabTemp.EOF Then
      SQL = "insert into ObservacaoNota (Codigo,Descricao,Observacao)"
      SQL = SQL & " values(900,'Tributos Totais NFC','Tributos Totais Incidentes(Lei Federal 12.741/2012) :')"
      CONECTA_GLOBAL.Execute SQL
   End If
   If TabTemp.State = 1 Then _
      TabTemp.Close
   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close
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
   If Not EXISTE_OBJ_BANCO("RETAGUARDA", "IMPREL", "U") = True Then
      SQL = "CREATE TABLE [dbo].[IMPREL]("
      SQL = SQL & " [IMPREL_ID] [bigint] NOT NULL,"
      SQL = SQL & " [RELATORIO] [nvarchar](60) NOT NULL,"
      SQL = SQL & " [CAMINHO] [nvarchar](max) NOT NULL,"
      SQL = SQL & " CONSTRAINT [PK_IMPREL] PRIMARY KEY CLUSTERED"
      SQL = SQL & " ([IMPREL_ID] Asc"
      SQL = SQL & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, "
      SQL = SQL & " IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) "
      SQL = SQL & " ON [PRIMARY]) ON [PRIMARY]"
      CONECTA_RETAGUARDA.Execute SQL
   End If

   If EXISTE_OBJ_BANCO("RETAGUARDA", "PEDIDOITEMOBS", "U") = False Then
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
   End If
   If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_PEDIDOITEMOBS_PEDIDOITEM", "") = False Then
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
   frmCADASTROPRODUTO.cmbSt.Visible = False
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
   frmCADASTROPRODUTO.txtFornec.Visible = False
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
PB.Max = Rs.RecordCount
PB.Value = 0
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
   PB.Value = i
   Print #1, "<TR>"
   For Each f In Rs.Fields
      Print #1, "<TD>" & Rs.Fields(f.Name) & "</TD>"
   Next
   Print #1, "</TR>"
   Rs.MoveNext
Next
Screen.MousePointer = vbNormal
MsgBox "Conversão realizada com sucesso"
PB.Value = 0
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
PB.Max = Rs.RecordCount
PB.Value = 0
Screen.MousePointer = vbHourglass
XMLRootTags = "<Root></Root>"
done = xmlDoc.loadXML(XMLRootTags)

If done = True Then
    Set xmlroot = xmlDoc.documentElement
    For i = 1 To Rs.RecordCount
        PB.Value = i
        Set node = xmlDoc.createNode(NODE_ELEMENT, "Node" & i, "")
        xmlroot.appendChild node
        Set nodeTag = xmlroot.selectSingleNode("Node" & i)
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
MsgBox "Conversão realizada com sucesso"
PB.Value = 0
xmlDoc.Save txtDestino.Text
Rs.Close
Set Rs = Nothing
cnn.Close
Set cnn = Nothing
' Exit Sub
'CheckXml:
' MsgBox Err.Description

End Sub

Public Function Mostra_Descrição_TipoVenda(TIPOVENDA_ID As Long) As String
'On Error GoTo ERRO_TRATA

   Dim TabTipovenda As New ADODB.Recordset

   If Not IsNull(TIPOVENDA_ID) Then
      SQL = "select descricao from TIPOVENDA WITH (NOLOCK)"
      SQL = SQL & " where tipovenda_id = " & TIPOVENDA_ID
      TabTipovenda.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabTipovenda.EOF Then _
         If Trim(TabTipovenda.Fields(0).Value) <> "" Then _
            Mostra_Descrição_TipoVenda = Trim(TabTipovenda.Fields(0).Value)
   End If
   If TabTipovenda.State = 1 Then _
      TabTipovenda.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "Mostra_Descrição_TipoVenda"
End Function

Public Function CONVERTE_VALOR_GRAMA(VALOR_VENDA As Double, VALOR_KILO As Double, PRODUTO_ID_N As Long) As Double
'On Error GoTo ERRO_TRATA

   Dim TabKilo As New ADODB.Recordset

   CONVERTE_VALOR_GRAMA = 0

   If PRODUTO_ID_N > 0 Then
      If TabKilo.State = 1 Then _
         TabKilo.Close

      SQL = "select preco_venda from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where produto_id = " & PRODUTO_ID_N
      TabKilo.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabKilo.EOF Then _
         If Not IsNull(TabKilo.Fields(0).Value) Then _
            VALOR_KILO = 0 & TabKilo.Fields(0).Value
      If TabKilo.State = 1 Then _
         TabKilo.Close
   End If
'no cadastro de produto o valor de venda é gravado baseado em um kilo
   If VALOR_VENDA > 0 And VALOR_KILO > 0 Then _
      CONVERTE_VALOR_GRAMA = VALOR_VENDA / VALOR_KILO

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "CONVERTE_VALOR_GRAMA"
End Function

Public Sub RODA_AT_ESTOQUE(PROD_ID_N As Long, Estab_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
   If PROD_ID_N > 0 Then _
      SQL = SQL & " where produto_id = " & PROD_ID_N
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabProduto.EOF
      DoEvents

      If TabTemp.State = 1 Then _
         TabTemp.Close

      SQL = "select * from ESTOQUE WITH (NOLOCK)"
      SQL = SQL & " where produto_id = " & TabProduto.Fields("produto_id").Value
      SQL = SQL & " and estabelecimento_id = " & Estab_ID_N
      TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabTemp.EOF Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "insert into ESTOQUE "
         SQL = SQL & " (ESTOQUE_ID,ESTABELECIMENTO_ID,PRODUTO_ID,QTDE_ESTOQUE)"
         SQL = SQL & " values("
            SQL = SQL & MAX_ID("estoque_id", "estoque", "", "", "", "") 'ESTOQUE_ID
            SQL = SQL & "," & Estab_ID_N                                'ESTABELECIMENTO_ID
            SQL = SQL & "," & TabProduto.Fields("produto_id").Value     'PRODUTO_ID
            SQL = SQL & "," & tpMOEDA(0)                                'QTDE_ESTOQUE
         SQL = SQL & ")"

         CONECTA_RETAGUARDA.Execute SQL
      End If
      If TabTemp.State = 1 Then _
         TabTemp.Close

      TabProduto.MoveNext
   Wend
   If TabProduto.State = 1 Then _
      TabProduto.Close

   QTDE_N = 0

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "RODA_AT_ESTOQUE"
End Sub

Public Sub ATUALIZA_ESTOQUE(NUMR_PRODUTO_ID_N As Long, NUMR_PEDIDO_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If NUMR_PEDIDO_ID_N <= 0 Then _
      Exit Sub

   Dim TabEstoqueAT As New ADODB.Recordset

   If TabEstoqueAT.State = 1 Then _
      TabEstoqueAT.Close

   SQL = "select PEDIDOITEM.PRODUTO_ID, PEDIDOITEM.QTD_PEDIDA,empresa_id,estabelecimento_id"
   SQL = SQL & " from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDOITEM WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID "
   SQL = SQL & " and PEDIDO.PEDIDO_ID = PEDIDOITEM.PEDIDO_ID"

   SQL = SQL & " where PEDIDO.pedido_id = " & NUMR_PEDIDO_ID_N
   SQL = SQL & " and PEDIDO.estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and pedidoitem.status <> 'C' "

   If NUMR_PRODUTO_ID_N > 0 Then _
      SQL = SQL & " and produto_id = " & NUMR_PRODUTO_ID_N

   TabEstoqueAT.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   While Not TabEstoqueAT.EOF
      RODA_AT_ESTOQUE TabEstoqueAT.Fields("produto_id").Value, TabEstoqueAT.Fields("estabelecimento_id").Value

'set colocar sp aqui

      '================estoque
      SQL = "update ESTOQUE set "

      SQL = SQL & " QTDE_ESTOQUE = QTDE_ESTOQUE - " & tpMOEDA(TabEstoqueAT.Fields("QTD_PEDIDA").Value)

      SQL = SQL & " from EMPRESA "
      SQL = SQL & " INNER JOIN ESTABELECIMENTO "
      SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID "
      SQL = SQL & " INNER JOIN ESTOQUE "
      SQL = SQL & " ON ESTABELECIMENTO.ESTABELECIMENTO_ID = ESTOQUE.ESTABELECIMENTO_ID"

      SQL = SQL & " where produto_id = " & TabEstoqueAT.Fields("produto_id").Value
      SQL = SQL & " and ESTABELECIMENTO.empresa_id = " & TabEstoqueAT.Fields("empresa_id").Value
      SQL = SQL & " and ESTOQUE.estabelecimento_id = " & TabEstoqueAT.Fields("estabelecimento_id").Value

      CONECTA_RETAGUARDA.Execute SQL
      '=======================

      SQL = "update produto set dt_ult_venda = '" & Now & "'"
      SQL = SQL & " where produto_id = " & TabEstoqueAT.Fields("produto_id").Value
      CONECTA_RETAGUARDA.Execute SQL

      TabEstoqueAT.MoveNext
   Wend
   If TabEstoqueAT.State = 1 Then _
      TabEstoqueAT.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "ATUALIZA_ESTOQUE"
End Sub

Public Function EXISTE_PRODUTO_CADASTRADO(CODG_PROD_A As String, Descricao_A As String) As Boolean
'On Error GoTo ERRO_TRATA

   Dim TabEstoque As New ADODB.Recordset

   EXISTE_PRODUTO_CADASTRADO = False

   If Trim(CODG_PROD_A) <> "" Or Trim(Descricao_A) <> "" Then
      If TabEstoque.State = 1 Then _
         TabEstoque.Close

      SQL = "select produto_id from PRODUTO WITH (NOLOCK)"
      SQL = SQL & " where estabelecimento_id = " & ESTAB_ID

      If Trim(CODG_PROD_A) <> "" Then _
         SQL = SQL & " and codg_produto = '" & Trim(CODG_PROD_A) & "'"
      If Trim(Descricao_A) <> "" Then _
         SQL = SQL & " and descricao = '" & Trim(Descricao_A) & "'"

      TabEstoque.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEstoque.EOF Then _
         If Not IsNull(TabEstoque.Fields(0).Value) Then _
            EXISTE_PRODUTO_CADASTRADO = True
      If TabEstoque.State = 1 Then _
         TabEstoque.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "EXISTE_PRODUTO_CADASTRADO"
End Function

Public Sub EXCLUIR_REGISTRO_FONE(Numero_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(Numero_A) <> "" And PESSOA_ID_N > 0 Then
      SQL = "delete FONE "
      SQL = SQL & " where numero = '" & Trim(Numero_A) & "'"
      SQL = SQL & " and pessoa_id = " & PESSOA_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "MDLGERAL", "EXCLUIR_REGISTRO_FONE"
End Sub

Public Sub GRAVA_ESTABELECIMENTOACESSO(USU_ID_N As Long, Estab_ID_N As Long)
'On Error GoTo ERRO_TRATA

   If TabConsulta.State = 1 Then _
      TabConsulta.Close

   SQL = "select * from ESTABELECIMENTOACESSO WITH (NOLOCK)"
   SQL = SQL & " where estabelecimento_id = " & Estab_ID_N
   SQL = SQL & " and usuario_id = " & USU_ID_N
   TabConsulta.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabConsulta.EOF Then
      SQL = "insert into ESTABELECIMENTOACESSO "
      SQL = SQL & "values("
         SQL = SQL & Estab_ID_N
         SQL = SQL & "," & USU_ID_N
      SQL = SQL & " )"
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabConsulta.State = 1 Then _
      TabConsulta.Close
      
Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "MDLGERAL", "GRAVA_ESTABELECIMENTOACESSO"
End Sub

'Public Enum ExtendedColorTypes
'    'vbWhite = &HFFFFFF
'    vbLightGray = &HE0E0E0
'    vbGray = &HC0C0C0
'    vbMediumGray = &H808080
'    vbDarkGray = &H404040
'    'vbBlack = &H0
'    vbPaleRed = &HC0C0FF
'    vbLightRed = &H8080FF
'    'vbRed = &HFF
'    vbMediumRed = &HC0&
'    vbDarkRed = &H80&
'    vbBlackRed = &H40&
'    vbPaleOrange = &HC0E0FF
'    vbLightOrange = &H80C0FF
'    vbOrange = &H80FF&
'    vbMediumOrange = &H40C0&
'    vbDarkOrange = &H4080&
'    vbBlackOrange = &H404080
'    vbPaleYellow = &HC0FFFF
'    vbLightYellow = &H80FFFF
'    'vbYellow = &HFFFF
'    vbMediumYellow = &HC0C0&
'    vbDarkYellow = &H8080&
'    vbBlackYellow = &H4040&
'    vbPaleGreen = &HC0FFC0
'    vbLightGreen = &H80FF80
'    'vbGreen = &HFF00
'    vbMediumGreen = &HC000&
'    vbDarkGreen = &H8000&
'    vbBlackGreen = &H4000&
'    vbPaleCyan = &HFFFFC0
'    vbLightCyan = &HFFFF80
'    'vbCyan = &HFFFF00
'    vbMediumCyan = &HC0C000
'    vbDarkCyan = &H808000
'    vbBlackCyan = &H404000
'    vbPaleBlue = &HFFC0C0
'    vbLightBlue = &HFF8080
'    'vbBlue = &HFF0000
'    vbMediumBlue = &HC00000
'    vbDarkBlue = &H800000
'    vbBlackBlue = &H400000
'    vbPalePurple = &HFFC0FF
'    vbLightPurple = &HFF80FF
'    vbPurple = &HFF00FF
'    'vbMagenta = &HFF00FF
'    vbMediumPurple = &HC000C0
'    vbDarkPurple = &H800080
'    vbBlackPurple = &H400040
'End Enum

Public Function TRAZ_VALOR_ITEMLANCAMENTO(ID_N As Long) As Double
'On Error GoTo ERRO_TRATA

   TRAZ_VALOR_ITEMLANCAMENTO = 0

   If ID_N > 0 Then
      Dim TabLocal   As New ADODB.Recordset

      If TabLocal.State = 1 Then _
        TabLocal.Close
      SQL = "select sum(valor_Item) as ValorTotal from ITEMLANCAMENTO "
      SQL = SQL & " where lancamento_id = " & ID_N
      TabLocal.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabLocal.EOF Then _
         If Not IsNull(TabLocal.Fields(0).Value) Then _
            TRAZ_VALOR_ITEMLANCAMENTO = 0 & TabLocal.Fields(0).Value
      If TabLocal.State = 1 Then _
        TabLocal.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "TRAZ_QTDE_ESTOQUE"
End Function

Public Function TRAZ_QTDE_ESTOQUE(ESTAB_ID As Long, PROD_ID_N As Long) As Double
'On Error GoTo ERRO_TRATA

   TRAZ_QTDE_ESTOQUE = 0
   If ESTAB_ID > 0 And PROD_ID_N >= 0 Then
      Dim TabEstoque As New ADODB.Recordset
      Dim strSQL  As String

      If TabEstoque.State = 1 Then _
         TabEstoque.Close

      strSQL = "select qtde_estoque from ESTOQUE WITH (NOLOCK)"
      strSQL = strSQL & " where estabelecimento_id = " & ESTAB_ID
      strSQL = strSQL & " and produto_id = " & PROD_ID_N
      TabEstoque.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
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

Public Function TRAZ_EMAIL(Pessoa_id As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_EMAIL = ""
   If Pessoa_id > 0 Then
      Dim TabEmail   As New ADODB.Recordset
      Dim strSQL  As String

      If TabEmail.State = 1 Then _
         TabEmail.Close

      strSQL = "select EMAIL from EMAIL WITH (NOLOCK)"
      strSQL = strSQL & " where pessoa_id = " & Pessoa_id
      TabEmail.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabEmail.EOF Then _
         TRAZ_EMAIL = "" & Trim(TabEmail.Fields("EMAIL").Value)
      If TabEmail.State = 1 Then _
         TabEmail.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_EMAIL"
End Function

Public Function TRAZ_RG(Pessoa_id As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_RG = ""
   If Pessoa_id > 0 Then
      Dim TabRG   As New ADODB.Recordset
      Dim strSQL  As String

      If TabRG.State = 1 Then _
         TabRG.Close

      strSQL = "select NUMERO_RG from RG WITH (NOLOCK)"
      strSQL = strSQL & " where pessoa_id = " & Pessoa_id
      TabRG.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabRG.EOF Then _
         TRAZ_RG = "" & Trim(TabRG.Fields("NUMERO_RG").Value)
      If TabRG.State = 1 Then _
         TabRG.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_RG"
End Function

Public Function TRAZ_IE(Pessoa_id As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_IE = ""
   If Pessoa_id > 0 Then
      Dim tabIE   As New ADODB.Recordset
      Dim strSQL  As String

      If tabIE.State = 1 Then _
         tabIE.Close

      strSQL = "select numr_ie from IE WITH (NOLOCK)"
      strSQL = strSQL & " where pessoa_id = " & Pessoa_id
      tabIE.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabIE.EOF Then _
         TRAZ_IE = "" & Trim(tabIE.Fields("numr_ie").Value)
      If tabIE.State = 1 Then _
         tabIE.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_IE"
End Function

Public Function TRAZ_IM(Pessoa_id As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_IM = ""
   If PESSOA_ID_N > 0 Then
      Dim tabIM   As New ADODB.Recordset
      Dim strSQL  As String

      If tabIM.State = 1 Then _
         tabIM.Close

      strSQL = "select numr_IM from IM WITH (NOLOCK)"
      strSQL = strSQL & " where pessoa_id = " & Pessoa_id
      tabIM.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabIM.EOF Then _
         TRAZ_IM = "" & Trim(tabIM.Fields("numr_IM").Value)
      If tabIM.State = 1 Then _
         tabIM.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_IM"
End Function

Public Function TRAZ_DADOS_NFe(TABELA_A As String, CAMPO_A As String, PESQUISA_A As String, DADO_A As String) As String
'On Error GoTo ERRO_TRATA

   TRAZ_DADOS_NFe = ""

   If Trim(TABELA_A) <> "" And Trim(CAMPO_A) <> "" And Trim(PESQUISA_A) <> "" And Trim(DADO_A) <> "" Then
      Dim TabTempGLOBAL As New ADODB.Recordset
      Dim strSQL        As String

      ABRE_BANCO_GLOBAL

      If CONECTA_GLOBAL.State = 1 Then
         If TabTempGLOBAL.State = 1 Then _
            TabTempGLOBAL.Close

         strSQL = "select " & Trim(CAMPO_A) & " from " & Trim(TABELA_A)
         strSQL = strSQL & " where " & Trim(PESQUISA_A) & " = '" & Trim(DADO_A) & "'"
         TabTempGLOBAL.Open strSQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabTempGLOBAL.EOF Then _
            If Not IsNull(TabTempGLOBAL.Fields(0).Value) Then _
               If Trim(TabTempGLOBAL.Fields(0).Value) <> "" Then _
                  TRAZ_DADOS_NFe = Trim(TabTempGLOBAL.Fields(0).Value)

         If TabTempGLOBAL.State = 1 Then _
            TabTempGLOBAL.Close

         If CONECTA_GLOBAL.State = 1 Then _
            CONECTA_GLOBAL.Close
      End If
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_DADOS_NFe"
End Function

Public Function TRAZ_CFOP(CFOP_ID_N As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_CFOP = ""
   If CFOP_ID_N > 0 Then
      Dim strSQL  As String

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      strSQL = "select descricao from CFOP WITH (NOLOCK)"
      strSQL = strSQL & " where cfop_id = " & CFOP_ID_N
      TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         TRAZ_CFOP = "" & Trim(TabDESCR.Fields("descricao").Value)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_CFOP"
End Function

Public Function TRAZ_CFOP_MSG(CFOP_ID_N As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_CFOP_MSG = ""
   If CFOP_ID_N > 0 Then
      Dim strSQL  As String

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      strSQL = "select msgfisco from CFOP WITH (NOLOCK)"
      strSQL = strSQL & " where cfop_id = " & CFOP_ID_N
      TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         TRAZ_CFOP_MSG = "" & Trim(TabDESCR.Fields("msgfisco").Value)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_CFOP_MSG"
End Function

Public Function TRAZ_DESCRITOR(TIPO As String, Codigo_N As String) As String
'On Error GoTo ERRO_TRATA

   TRAZ_DESCRITOR = ""
   If Trim(TIPO) <> "" And Codigo_N <> "" Then
      Dim strSQL  As String

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      strSQL = "select DESCRICAO from DESCR WITH (NOLOCK)"
      strSQL = strSQL & " where codigo <> '' "

      If IsNumeric(Codigo_N) Then _
         strSQL = strSQL & " and codigo = '" & Trim(Codigo_N) & "'"

      If Trim(TIPO) <> "" Then _
         strSQL = strSQL & " and TIPO = '" & Trim(TIPO) & "'"

      TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         TRAZ_DESCRITOR = "" & Trim(TabDESCR.Fields("DESCRICAO").Value)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_DESCRITOR"
End Function

Public Function TRAZ_ESTABELECIMENTO(ESTAB_ID As Integer) As String
'On Error GoTo ERRO_TRATA

   TRAZ_ESTABELECIMENTO = ""

   If Not IsNull(ESTAB_ID) Then
      Dim strSQL  As String

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      strSQL = "select descricao from ESTABELECIMENTO WITH (NOLOCK)"
      strSQL = strSQL & " where estabelecimento_id = " & ESTAB_ID
      TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         If Not IsNull(TabDESCR.Fields(0).Value) Then _
            TRAZ_ESTABELECIMENTO = Trim(TabDESCR.Fields(0).Value)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_ESTABELECIMENTO"
End Function

Public Function TRAZ_TIPO_USUARIO() As Long
'On Error GoTo ERRO_TRATA

   TRAZ_TIPO_USUARIO = 0

   If Not IsNull(USUARIO_ID_N) Then
      Dim strSQL  As String

      If TabDESCR.State = 1 Then _
         TabDESCR.Close
'1  OPERADOR
'2  VENDEDOR (a)
'3  ADMINISTRATIVO
'4  GERENTE
'5  DIRETOR
'6  FINANCEIRO
'7  CAIXA
      strSQL = "select tipo from USUARIO WITH (NOLOCK)"
      strSQL = strSQL & " where usuario_id = " & USUARIO_ID_N
      TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         If Not IsNull(TabDESCR.Fields(0).Value) Then _
            TRAZ_TIPO_USUARIO = TabDESCR.Fields(0).Value
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_TIPO_USUARIO"
End Function

Public Function TRAZ_NOME_USUARIO(USU_ID_N As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_NOME_USUARIO = ""

   If IsNull(USU_ID_N) Then _
      Exit Function

   If USU_ID_N < 0 Then _
      Exit Function

   Dim strSQL  As String

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   strSQL = "select nome from USUARIO WITH (NOLOCK)"
   strSQL = strSQL & " where usuario_id = " & USU_ID_N
   TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then _
      If Not IsNull(TabDESCR.Fields(0).Value) Then _
         TRAZ_NOME_USUARIO = Trim(TabDESCR.Fields(0).Value)
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_NOME_USUARIO"
End Function

Public Function TRAZ_NOME_PESSOA(Pessoa_id As Long, CNPJCPF As String) As String
'On Error GoTo ERRO_TRATA

   TRAZ_NOME_PESSOA = ""

   If Pessoa_id <= 0 And Trim(CNPJCPF) = "" Then _
      Exit Function

   Dim strSQL  As String
   Dim TabNome As New ADODB.Recordset

   If TabNome.State = 1 Then _
      TabNome.Close

   strSQL = "select descricao from PESSOA WITH (NOLOCK)"

   If Pessoa_id > 0 Then
      strSQL = strSQL & " where pessoa_id = " & Pessoa_id
      Else: strSQL = strSQL & " where cnpjcpf = '" & Trim(CNPJCPF) & "'"
   End If

   TabNome.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNome.EOF Then _
      If Not IsNull(TabNome.Fields(0).Value) Then _
         TRAZ_NOME_PESSOA = "" & Trim(TabNome.Fields(0).Value)
   If TabNome.State = 1 Then _
      TabNome.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_NOME_USUARIO"
End Function

Public Function TRAZ_NOME_FORNECEDOR(FORNEC_N As Long, PESSOA_N As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_NOME_FORNECEDOR = ""

   If IsNull(PESSOA_N) Then _
      Exit Function
   If PESSOA_N <= 0 Then _
      Exit Function

   If IsNull(FORNEC_N) Then _
      Exit Function
   If FORNEC_N <= 0 Then _
      Exit Function

   Dim strSQL  As String
   Dim TabNome As New ADODB.Recordset

   If TabNome.State = 1 Then _
      TabNome.Close

   strSQL = "select PESSOA.DESCRICAO from PESSOA WITH (NOLOCK)"
   strSQL = strSQL & " INNER JOIN FORNECEDOR  WITH (NOLOCK)"
   strSQL = strSQL & " ON PESSOA.PESSOA_ID = FORNECEDOR.PESSOA_ID"

   strSQL = strSQL & " where PESSOA.pessoa_id = " & PESSOA_N
   strSQL = strSQL & " and fornecedor_id = " & FORNEC_N
   TabNome.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNome.EOF Then _
      If Not IsNull(TabNome.Fields(0).Value) Then _
         TRAZ_NOME_FORNECEDOR = "" & Trim(TabNome.Fields(0).Value)
   If TabNome.State = 1 Then _
      TabNome.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_NOME_FORNECEDOR"
End Function

Public Function TRAZ_NOME_VENDEDOR(VEND_ID_N As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_NOMVE_VENDEDOR = ""

   If IsNull(VEND_ID_N) Then _
      Exit Function

   If Trim(VEND_ID_N) < 0 Then _
      Exit Function

   Dim strSQL  As String

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   strSQL = "SELECT PESSOA.PESSOA_ID, PESSOA.CNPJCPF, PESSOA.DESCRICAO, VENDEDOR.VENDEDOR_ID, "
   strSQL = strSQL & " VENDEDOR.EQUIPE_ID, VENDEDOR.STATUS, VENDEDOR.TABELAPRECO_ID"
   strSQL = strSQL & " FROM PESSOA WITH (NOLOCK)"
   strSQL = strSQL & " INNER JOIN VENDEDOR WITH (NOLOCK)"
   strSQL = strSQL & " ON PESSOA.PESSOA_ID = VENDEDOR.PESSOA_ID"
   strSQL = strSQL & " where vendedor_id = " & VEND_ID_N
   TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then _
      If Not IsNull(TabDESCR.Fields(0).Value) Then _
         TRAZ_NOMVE_VENDEDOR = Trim(TabDESCR.Fields(0).Value)
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_NOME_VENDEDOR"
End Function

Public Function TRAZ_DESCRICAO_PRODUTO(PROD_ID_A As String, CODG_PROD_A As String) As String
'On Error GoTo ERRO_TRATA

   TRAZ_DESCRICAO_PRODUTO = ""

   If IsNull(PROD_ID_A) And IsNull(CODG_PROD_A) Then _
      Exit Function

   If Trim(PROD_ID_A) = "" And Trim(CODG_PROD_A) = "" Then _
      Exit Function

   Dim strSQL  As String

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   strSQL = "select descricao from PRODUTO WITH (NOLOCK)"
   If Trim(CODG_PROD_A) <> "" Then
       strSQL = strSQL & " where codg_produto = '" & Trim(CODG_PROD_A) & "'"
      Else: strSQL = strSQL & " where produto_id = " & PROD_ID_A
   End If
   TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then _
      If Not IsNull(TabDESCR.Fields(0).Value) Then _
         TRAZ_DESCRICAO_PRODUTO = Trim(TabDESCR.Fields(0).Value)
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_DESCRICAO_PRODUTO"
End Function

Public Function TRAZ_DESCRICAO_FORMAPAGTO(FORMA_ID_N As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_DESCRICAO_FORMAPAGTO = ""

   If IsNull(FORMA_ID_N) Then _
      Exit Function

   If Trim(FORMA_ID_N) < 0 Then _
      Exit Function

   Dim strSQL  As String

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   strSQL = "select descricao from FORMAPAGTO WITH (NOLOCK)"
   strSQL = strSQL & " where formapagto_id = " & FORMA_ID_N
   TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then _
      If Not IsNull(TabDESCR.Fields(0).Value) Then _
         TRAZ_DESCRICAO_FORMAPAGTO = Trim(TabDESCR.Fields(0).Value)
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_DESCRICAO_FORMAPAGTO"
End Function

Public Function TRAZ_DESCRICAO_TIPOVENDA(TPVENDA_ID_N As Long) As String
'On Error GoTo ERRO_TRATA

   TRAZ_DESCRICAO_TIPOVENDA = ""

   If IsNull(TPVENDA_ID_N) Then _
      Exit Function

   If Trim(TPVENDA_ID_N) < 0 Then _
      Exit Function

   Dim strSQL  As String

   If TabDESCR.State = 1 Then _
      TabDESCR.Close

   strSQL = "select descricao from TIPOVENDA WITH (NOLOCK)"
   strSQL = strSQL & " where TIPOVENDA_id = " & TPVENDA_ID_N
   TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabDESCR.EOF Then _
      If Not IsNull(TabDESCR.Fields(0).Value) Then _
         TRAZ_DESCRICAO_TIPOVENDA = Trim(TabDESCR.Fields(0).Value)
   If TabDESCR.State = 1 Then _
      TabDESCR.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_DESCRICAO_TIPOVENDA"
End Function

Public Function TRAZ_ID_TABELA(NOME_TABELA As String, NOME_CAMPO As String, Campo1_A As String, Condicao1_A As String) As Long
'On Error GoTo ERRO_TRATA

   TRAZ_ID_TABELA = 0
   If Trim(NOME_TABELA) <> "" And Trim(NOME_CAMPO) <> "" And Trim(Campo1_A) <> "" And Trim(Condicao1_A) <> "" Then
      Dim strSQL  As String

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      strSQL = "select " & NOME_CAMPO & " from  " & NOME_TABELA
      strSQL = strSQL & " where  " & Campo1_A & " = '" & Condicao1_A & "'"
      TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         If Not IsNull(TabDESCR.Fields(0).Value) Then _
            TRAZ_ID_TABELA = 0 & Trim(TabDESCR.Fields(0).Value)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_ID_TABELA"
End Function

Public Function TRAZ_TEXTO_TABELA(NOME_TABELA As String, NOME_CAMPO As String, Campo1_A As String, Condicao1_A As String) As String
'On Error GoTo ERRO_TRATA

   TRAZ_TEXTO_TABELA = 0
   If Trim(NOME_TABELA) <> "" And Trim(NOME_CAMPO) <> "" And Trim(Campo1_A) <> "" And Trim(Condicao1_A) <> "" Then
      Dim strSQL  As String

      If TabDESCR.State = 1 Then _
         TabDESCR.Close

      strSQL = "select " & NOME_CAMPO & " from  " & NOME_TABELA
      strSQL = strSQL & " where  " & Campo1_A & " = '" & Condicao1_A & "'"
      TabDESCR.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabDESCR.EOF Then _
         If Not IsNull(TabDESCR.Fields(0).Value) Then _
            TRAZ_TEXTO_TABELA = "" & Trim(TabDESCR.Fields(0).Value)
      If TabDESCR.State = 1 Then _
         TabDESCR.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_TEXTO_TABELA"
End Function

Public Function TRAZ_PRECO_VENDA_PRODUTO_TABPRECO(PROD_N As Long, TAB_PRECO_ID As Integer, FORMAPAGTO_ID As Integer) As Double
'On Error GoTo ERRO_TRATA

   TRAZ_PRECO_VENDA_PRODUTO_TABPRECO = 0
   If TAB_PRECO_ID > 0 And PROD_N > 0 Then
      Dim TabPreco As New ADODB.Recordset
      Dim strSQL  As String

      If TabPreco.State = 1 Then _
         TabPreco.Close

      strSQL = "select VALOR_VENDA from TABELAPRECO WITH (NOLOCK)"
      strSQL = strSQL & " INNER JOIN TABELAPRECOITEM "
      strSQL = strSQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID"

      strSQL = strSQL & " where produto_id = " & PROD_N
      strSQL = strSQL & " and TABELAPRECO.tabelapreco_id = " & TAB_PRECO_ID
      strSQL = strSQL & " and FORMAPAGTO_ID = " & FORMAPAGTO_ID
      strSQL = strSQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabPreco.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPreco.EOF Then _
         TRAZ_PRECO_VENDA_PRODUTO_TABPRECO = 0 & TabPreco.Fields(0).Value
      If TabPreco.State = 1 Then _
         TabPreco.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_PRECO_VENDA_PRODUTO_TABPRECO"
End Function

Public Function TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO(PROD_N As Long, TAB_PRECO_ID As Integer, FORMAPAGTO_ID As Integer) As Double
'On Error GoTo ERRO_TRATA

   TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO = 0
   If TAB_PRECO_ID > 0 And PROD_N > 0 Then
      Dim TabPreco As New ADODB.Recordset
      Dim strSQL  As String

      If TabPreco.State = 1 Then _
         TabPreco.Close

      strSQL = "select VALOR_CUSTO from TABELAPRECO WITH (NOLOCK)"
      strSQL = strSQL & " INNER JOIN TABELAPRECOITEM "
      strSQL = strSQL & " ON TABELAPRECO.TABELAPRECO_ID = TABELAPRECOITEM.TABELAPRECO_ID"

      strSQL = strSQL & " where produto_id = " & PROD_N
      strSQL = strSQL & " and TABELAPRECO.tabelapreco_id = " & TAB_PRECO_ID
      strSQL = strSQL & " and FORMAPAGTO_ID = " & FORMAPAGTO_ID
      strSQL = strSQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

      TabPreco.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabPreco.EOF Then _
         TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO = 0 & TabPreco.Fields(0).Value
      If TabPreco.State = 1 Then _
         TabPreco.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_PRECO_CUSTO_PRODUTO_TABPRECO"
End Function

Public Function TRAZ_NUMERO_CUPOM() As Long
'On Error GoTo ERRO_TRATA

   TRAZ_NUMERO_CUPOM = 0

   If PEDIDO_ID_N > 0 Then
      Dim TabECF  As New ADODB.Recordset
      Dim strSQL  As String

      If TabECF.State = 1 Then _
         TabECF.Close

      strSQL = "select numr_cupom from CUPOM WITH (NOLOCK)"
      strSQL = strSQL & " where pedido_id = " & PEDIDO_ID_N
      TabECF.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabECF.EOF Then _
         If Not IsNull(TabECF.Fields(0).Value) Then _
            TRAZ_NUMERO_CUPOM = TabECF.Fields(0).Value
      If TabECF.State = 1 Then _
         TabECF.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_NUMERO_CUPOM"
End Function
'===========================
Public Sub PEGA_DADOS_EMPRESA()
'On Error GoTo ERRO_TRATA

   Dim tabEmpresa As New ADODB.Recordset

   ENDERECO_ID_N = 0
   PESSOA_ID_EMPRESA_N = 0

   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

   SQL = "select EMPRESA.EMPRESA_ID, EMPRESA.CGC, EMPRESA.RAZAO_SOCIAL, "
   SQL = SQL & " EMPRESA.CFOP_DV_SAI_FE, EMPRESA.CFOP_DV_SAI_DE, EMPRESA.CFOP_DV_ENT_FE, "
   SQL = SQL & " EMPRESA.CFOP_DV_ENT_DE, EMPRESA.CFOP_TRA_SAI_FE, EMPRESA.CFOP_TRA_SAI_DE, "
   SQL = SQL & " EMPRESA.CFOP_TRA_ENT_FE, EMPRESA.CFOP_TRA_ENT_DE, EMPRESA.CFOP_SAIDA_FE,"
   SQL = SQL & " EMPRESA.CFOP_SAIDA_DE, EMPRESA.CFOP_ENTRADA_FE, EMPRESA.CFOP_ENTRADA_DE, "
   SQL = SQL & " EMPRESA.CFOP_VENDA_FORA_PAIS, EMPRESA.TP2_DE_CONTRIB, EMPRESA.TP2_DE_NCONTRIB,"
   SQL = SQL & " EMPRESA.TP2_DE_CMAQ_IMP, EMPRESA.TP2_DE_NMAQ_IMP, EMPRESA.TP2_FE_CMAQ_IMP, "
   SQL = SQL & " EMPRESA.TP2_FE_NMAQ_IMP, EMPRESA.TP2_FE_CAP_INDU, EMPRESA.TP2_FE_NAP_INDU,"
   SQL = SQL & " EMPRESA.TIPO_REGIME_EMPRESA, EMPRESA.TIPO_ENQUADRAMENTO_SIMPLES, EMPRESA.PESSOA_ID, "
   SQL = SQL & " EMPRESA.CRT, ENDERECO.CEP_ID, ENDERECO.RUA, ENDERECO.BAIRRO,"
   SQL = SQL & " Endereco.Complemento , Endereco.Numero, Empresa.SERIE_NOTA_SAIDA, Endereco.ENDERECO_ID, "
   SQL = SQL & " CEP.Cidade, CEP.UF, CEP.IBGE_ID"

   SQL = SQL & " from CEP WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ENDERECO WITH (NOLOCK) "
   SQL = SQL & " ON CEP.Cep_ID = ENDERECO.CEP_ID "
   SQL = SQL & " RIGHT OUTER JOIN EMPRESA WITH (NOLOCK) "
   SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (NOLOCK) "
   SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID "
   SQL = SQL & " ON ENDERECO.PESSOA_ID = EMPRESA.PESSOA_ID"

   SQL = SQL & " where EMPRESA.empresa_id = " & EMPRESA_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N

   SQL = SQL & " and cgc = '" & Trim(CNPJ_EMPRESA_N) & "'"
   SQL = SQL & " and tipo = 'C'" 'endereço comercial

   tabEmpresa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If tabEmpresa.EOF Then
      If tabEmpresa.State = 1 Then _
         tabEmpresa.Close

      MsgBox "O sistema não obteve sucesso ao tentar localizar a empresa corrente."
      Exit Sub
   End If

   PESSOA_ID_EMPRESA_N = 0 & tabEmpresa.Fields("PESSOA_ID").Value
   SERIE_NFe_A = "" & Trim(tabEmpresa.Fields("SERIE_NOTA_SAIDA").Value)
   TIPO_REGIME_EMPRESA_A = "" & Trim(tabEmpresa.Fields("TIPO_REGIME_EMPRESA").Value)
   ENDERECO_ID_N = 0 & Trim(tabEmpresa.Fields("ENDERECO_ID").Value)
   CCE_EMPRESA_N = "" & Trim(TRAZ_IE(PESSOA_ID_EMPRESA_N))
   UF_EMPRESA_A = "" & Trim(tabEmpresa.Fields("uf").Value)
   CTR_EMPRESA_N = Trim(tabEmpresa.Fields("CRT").Value)

   ' yuri 01/05/2012 para pegar tambem outras informações referentes a importos
   'g_trabalhacomtare_empresa = tabEmpresa!optante_tare não retirar sergio vamos precisar
   'so to colocando aqui com comentário para nao te atrapalhar

   TP2_DE_CONTRIB = "" & Trim(tabEmpresa!TP2_DE_CONTRIB)
   TP2_DE_NCONTRIB = "" & Trim(tabEmpresa!TP2_DE_NCONTRIB)
   TP2_DE_CMAQ_IMP = "" & Trim(tabEmpresa!TP2_DE_CMAQ_IMP)
   TP2_DE_NMAQ_IMP = "" & Trim(tabEmpresa!TP2_DE_NMAQ_IMP)
   TP2_FE_CMAQ_IMP = "" & Trim(tabEmpresa!TP2_FE_CMAQ_IMP)
   TP2_FE_NMAQ_IMP = "" & Trim(tabEmpresa!TP2_FE_NMAQ_IMP)
   TP2_FE_CAP_INDU = "" & Trim(tabEmpresa!TP2_FE_CAP_INDU)
   TP2_FE_NAP_INDU = "" & Trim(tabEmpresa!TP2_FE_NAP_INDU)

   CFOP_SAIDA_DENTRO_UF_N = "" & Trim(tabEmpresa!CFOP_SAIDA_DE)
   CFOP_SAIDA_FORA_UF_N = "" & Trim(tabEmpresa!CFOP_SAIDA_FE)

   CFOP_DEVOLUCAO_SAI_DENTRO_UF_N = "" & Trim(tabEmpresa.Fields("CFOP_DV_SAI_DE").Value)
   CFOP_DEVOLUCAO_SAI_FORA_UF_N = "" & Trim(tabEmpresa.Fields("CFOP_DV_SAI_FE").Value)

   'ENTRADA
   CFOP_DEVOLUCAO_ENTRADA_FORA_UF_N = "" & Trim(tabEmpresa.Fields("CFOP_DV_ENT_FE").Value)
   CFOP_DEVOLUCAO_ENTRADA_DENTRO_UF_N = "" & Trim(tabEmpresa.Fields("CFOP_DV_ENT_DE").Value)
   CFOP_ENTRADA_FE = "" & Trim(tabEmpresa.Fields("CFOP_ENTRADA_FE").Value)
   CFOP_ENTRADA_DE = "" & Trim(tabEmpresa.Fields("CFOP_ENTRADA_dE").Value)

   If tabEmpresa.State = 1 Then _
      tabEmpresa.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "PEGA_DADOS_EMPRESA"
End Sub
'========================================
Public Sub VERIFICA_TABELA_CLIENTE()
'On Error GoTo ERRO_TRATA

   If EXISTE_OBJ_BANCO("RETAGUARDA", "CLIENTE", "U") = True Then
      Dim VACA_VEIA_A As String
      VACA_VEIA_A = ""

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PROFISSAO", "CLIENTE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE DROP COLUMN PROFISSAO"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "codg_distribuidor", "CLIENTE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE DROP COLUMN codg_distribuidor"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODG_SUFRAMA", "CLIENTE") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE ADD CODG_SUFRAMA BIGINT"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "PESSOA_ID", "CLIENTE") = False Then
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE ADD PESSOA_ID BIGINT"
         Else: Alteração_Definição_Campo_Tabela "PESSOA_ID", "BIGINT", "CLIENTE", "RETAGUARDA"
      End If

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VENDEDOR", "CLIENTE") = True Then _
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CLIENTE.VENDEDOR'" & "," & "'VENDEDOR_ID'" & "," & "'COLUMN'"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "VENDEDOR_ID", "CLIENTE") = True Then _
         Alteração_Definição_Campo_Tabela "VENDEDOR_ID", "BIGINT", "CLIENTE", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CLIENTE_ID", "CLIENTE") = True Then _
         Alteração_Definição_Campo_Tabela "CLIENTE_ID", "BIGINT", "CLIENTE", "RETAGUARDA"

      If EXISTE_CAMPO_TABELA("RETAGUARDA", "CODIGO", "CLIENTE") = True Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE DROP COLUMN CODIGO"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "pk_CLIENTE", "") = False Then _
         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE ADD CONSTRAINT pk_CLIENTE PRIMARY KEY (CLIENTE_ID)"

      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CLIENTE_PESSOA", "") = False Then
         SQL = "ALTER TABLE [dbo].[CLIENTE]  WITH CHECK ADD  CONSTRAINT [FK_CLIENTE_PESSOA] FOREIGN KEY([PESSOA_ID])"
         SQL = SQL & " References [dbo].[PESSOA]([PESSOA_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[CLIENTE] CHECK CONSTRAINT [FK_CLIENTE_PESSOA]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
'=======================ALTERANDO O LINK PARA ESTABELECIMENTO
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CLIENTE_EMPRESA", "") = True Then
         SQL = "alter table CLIENTE "
         SQL = SQL & " drop CONSTRAINT FK_CLIENTE_EMPRESA"
         CONECTA_RETAGUARDA.Execute SQL

         'SQL = "ALTER TABLE [dbo].[CLIENTE]  WITH CHECK ADD  CONSTRAINT [FK_CLIENTE_EMPRESA] FOREIGN KEY([EMPRESA_ID])"
         'SQL = SQL & " References [dbo].[EMPRESA]([EMPRESA_ID])"
         'CONECTA_RETAGUARDA.Execute SQL
   
         'SQL = "ALTER TABLE [dbo].[CLIENTE] CHECK CONSTRAINT [FK_CLIENTE_EMPRESA]"
         'CONECTA_RETAGUARDA.Execute SQL
      End If
'=======================
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "EMPRESA_ID", "CLIENTE") = True Then
         CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'CLIENTE.EMPRESA_ID'" & "," & "'ESTABELECIMENTO_ID'" & "," & "'COLUMN'"
      
         SQL = "update cliente set cliente.ESTABELECIMENTO_ID = pedido.ESTABELECIMENTO_ID"
         SQL = SQL & " from PEDIDO"
         SQL = SQL & " INNER JOIN CLIENTE"
         SQL = SQL & " ON PEDIDO.CLIENTE_ID = CLIENTE.CLIENTE_ID"
         CONECTA_RETAGUARDA.Execute SQL
      End If
      If EXISTE_OBJ_BANCO("RETAGUARDA", "FK_CLIENTE_ESTABELECIMENTO", "") = False Then
         SQL = "ALTER TABLE [dbo].[CLIENTE]  WITH CHECK ADD  CONSTRAINT [FK_CLIENTE_ESTABELECIMENTO] FOREIGN KEY([ESTABELECIMENTO_ID])"
         SQL = SQL & " References [dbo].[ESTABELECIMENTO]([ESTABELECIMENTO_ID])"
         CONECTA_RETAGUARDA.Execute SQL
   
         SQL = "ALTER TABLE [dbo].[CLIENTE] CHECK CONSTRAINT [FK_CLIENTE_ESTABELECIMENTO]"
         CONECTA_RETAGUARDA.Execute SQL
      End If
'=========================== IE
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "IE", "CLIENTE") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where ie is not null"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            If Not IsNull(TabTemp.Fields("ie").Value) Then
               If IsNumeric(TabTemp.Fields("ie").Value) Then
                  PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value
                  ENDERECO_ID_N = 0 & TRAZ_ID_ENDERECO("C")
                  VACA_VEIA_A = ""
                  VACA_VEIA_A = "" & Replace(TabTemp.Fields("ie").Value, "-", "")
                  VACA_VEIA_A = Replace(VACA_VEIA_A, ".", "")
                  VACA_VEIA_A = Replace(VACA_VEIA_A, "/", "")
                  VACA_VEIA_A = Replace(VACA_VEIA_A, "\", "")
                  SQL = ""

                  GRAVA_IE VACA_VEIA_A
               End If
            End If

            PESSOA_ID_N = 0
            ENDERECO_ID_N = 0
            SQL = ""

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE DROP COLUMN IE"
      End If
      SQL = ""
'=========================== IM
      If EXISTE_CAMPO_TABELA("RETAGUARDA", "IM", "CLIENTE") = True Then
         If TabTemp.State = 1 Then _
            TabTemp.Close

         SQL = "select * from CLIENTE WITH (NOLOCK)"
         SQL = SQL & " where IM is not null"
         TabTemp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         While Not TabTemp.EOF
            If Not IsNull(TabTemp.Fields("IM").Value) Then
               If IsNumeric(TabTemp.Fields("IM").Value) Then
                  PESSOA_ID_N = 0 & TabTemp.Fields("pessoa_id").Value
                  ENDERECO_ID_N = 0 & TRAZ_ID_ENDERECO("C")
                  VACA_VEIA_A = ""
                  VACA_VEIA_A = "" & Replace(TabTemp.Fields("IM").Value, "-", "")
                  VACA_VEIA_A = Replace(VACA_VEIA_A, ".", "")
                  VACA_VEIA_A = Replace(VACA_VEIA_A, "/", "")
                  VACA_VEIA_A = Replace(VACA_VEIA_A, "\", "")
                  SQL = ""

                  GRAVA_IM VACA_VEIA_A
               End If
            End If

            PESSOA_ID_N = 0
            ENDERECO_ID_N = 0
            SQL = ""

            TabTemp.MoveNext
         Wend
         If TabTemp.State = 1 Then _
            TabTemp.Close

         CONECTA_RETAGUARDA.Execute "ALTER TABLE CLIENTE DROP COLUMN IM"
      End If
      SQL = ""
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "VERIFICA_TABELA_CLIENTE"
End Sub

Public Function TRAZ_ID_ENDERECO(Tipo_End_A As String) As Long
'On Error GoTo ERRO_TRATA

   TRAZ_ID_ENDERECO = 0

   If PESSOA_ID_N <= 0 Then _
      Exit Function
   If Tipo_End_A = "" Then _
      Exit Function

   Dim TabVai  As New ADODB.Recordset
   Dim strSQL  As String

   If TabVai.State = 1 Then _
      TabVai.Close
   SQL = "select endereco_id from ENDERECO WITH (NOLOCK)"
   SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
   SQL = SQL & " and tipo = '" & Trim(Tipo_End_A) & "'"
   TabVai.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabVai.EOF Then _
      If Not IsNull(TabVai.Fields(0).Value) Then _
         TRAZ_ID_ENDERECO = 0 & TabVai.Fields(0).Value
   If TabVai.State = 1 Then _
      TabVai.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "TRAZ_ID_ENDERECO"
End Function

Public Function TRATA_PESSOA(CNPJ_CPF_A As String) As Boolean
'On Error GoTo ERRO_TRATA

   TRATA_PESSOA = False

   Dim TabPessoa        As New ADODB.Recordset
   Dim rstAux           As New ADODB.Recordset
   Dim tabEndereco      As New ADODB.Recordset
   Dim rstCep           As New ADODB.Recordset
   Dim DIAS_DE_ATRAZO   As Integer

   ENDERECO_A = ""
   PESSOA_ID_N = 0
   VALOR_PENDENTE_N = 0
   If Trim(CNPJ_CPF_A) <> "" Then
      If CHECA_CNPJCPF(Trim(CNPJ_CPF_A)) = False Then
         MsgBox "CNPJ/CPF com DV incorreto !!! "
         CNPJ_CPF_A = ""
         Exit Function
      End If
   End If

   If TabPessoa.State = 1 Then _
      TabPessoa.Close

   SQL = "select * from PESSOA WITH (NOLOCK)"
   SQL = SQL & " where cnpjcpf = '" & Trim(CNPJ_CPF_A) & "'"
   TabPessoa.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPessoa.EOF Then
      If TabPessoa!SITUACAO = "C" Then
         If TabPessoa.State = 1 Then _
            TabPessoa.Close

         Beep
         MsgBox "Cliente Esta Bloqueado!, Verifique Cadastro!.", vbOKOnly, "Atenção."
         CNPJ_CPF_A = ""
         PESSOA_ID_N = 0
         Exit Function
      End If

      PESSOA_ID_N = 0 & TabPessoa.Fields("pessoa_id").Value
      NOME_CLIENTE_A = "" & TabPessoa.Fields("descricao").Value

      If TabCliente.State = 1 Then _
         TabCliente.Close
      SQL = "select limite_credito,cliente_id from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         LIMITE_CREDITO_CLI_N = 0 & TabCliente.Fields("LIMITE_CREDITO").Value
         CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close

      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      SQL = "select CLIENTE.CLIENTE_ID, ENDERECO.ENDERECO_ID, ENDERECO.PESSOA_ID, "
      SQL = SQL & " ENDERECO.CEP_ID, ENDERECO.RUA, ENDERECO.BAIRRO, ENDERECO.COMPLEMENTO, "
      SQL = SQL & " ENDERECO.TIPO, ENDERECO.NUMERO, CEP.Cidade , CEP.UF, CEP.IBGE_ID"
      SQL = SQL & " from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " INNER JOIN ENDERECO WITH (NOLOCK)"
      SQL = SQL & " ON CLIENTE.PESSOA_ID = ENDERECO.PESSOA_ID "
      SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
      SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"

      SQL = SQL & " where CLIENTE.pessoa_id = " & PESSOA_ID_N
      SQL = SQL & " and tipo = 'C'"

      tabEndereco.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not tabEndereco.EOF Then
         ENDERECO_A = "" & Trim(tabEndereco!Rua)
         ENDERECO_A = ENDERECO_A & "," & "" & Trim(tabEndereco!Complemento)
         ENDERECO_A = ENDERECO_A & "," & "" & Trim(tabEndereco!Bairro)
         UF_CLIENTE_A = Trim(tabEndereco!UF)

         If USA_NFe = True And INDR_CAIXA = False Then
            'Pegou o CEP do cliente
            If IsNull(tabEndereco!CEP_ID) Then
               If tabEndereco.State = 1 Then _
                  tabEndereco.Close

               MsgBox "O Cadastro do cliente não está completo. Verique os dados (CEP_id, UF, Endereço, Insc. Estadual etc..)" & vbCrLf & "O sitema não pode continuar sem o devido acerto.", vbCritical
               CNPJ_CPF_A = ""
               Exit Function
            End If
         End If
         'Else: MsgBox "NÃO ACHOU CADASTRO DO CLIENTE ROTINA (TRATA_PESSOA)"
      End If
      If tabEndereco.State = 1 Then _
         tabEndereco.Close

      '=====================================
      'Pegou o tipo do cliente
      'PARECE QUE TIPO_CLIENTE É PRA DEFINIR SE ELE É PJ OU PF, ENTÃO VOU MUDAR A REGRA PARA
      'QUE SE ELE TEM INSCRIÇÃO ESTADUAL REALIZA A TRIBUTAÇÃO, CASO CONTRÁRIO, SE FOR PESSOA FISICA
      'NÃO DESTACA NA NOTA FISCAL.
      'TIPO_CLIENTE_N = 1
      'If Not IsNull(TabPessoa!TIPO_CLIENTE) Then _
         TIPO_CLIENTE_N = TabPessoa!TIPO_CLIENTE
      '=====================================

      CCE_CLIENTE_A = "ISENTO"
      CCE_CLIENTE_A = Trim(TRAZ_IE(PESSOA_ID_N))

      If Trim(CCE_CLIENTE_A) <> "ISENTO" Then
         TIPO_CLIENTE_N = 1
         If Trim(CCE_CLIENTE_A) = "0" Or Trim(CCE_CLIENTE_A) = "" Then
            Else
               If Valida_Inscricao_Estadual(CCE_CLIENTE_A, UF_CLIENTE_A) <> 0 Then
                  TIPO_CLIENTE_N = 2
               End If
         End If
      End If
      '==================================
      If INDR_PEDIDO_VALIDO = True Then
         SQL = "update PEDIDO set "
         SQL = SQL & " nome_cliente = '" & Trim(NOME_CLIENTE_A) & "'"
         SQL = SQL & ", cgccpf = '" & Trim(CNPJ_CPF_A) & "'"
         SQL = SQL & ", cliente_id = " & CLIENTE_ID_N
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
         CONECTA_RETAGUARDA.Execute SQL
      End If

      If Trim(CNPJ_CPF_A) <> "99999999999" Then
         If USA_NFe = True Then
            If LIMITE_CREDITO_CLI_N > 0 Then
               If rstAux.State = 1 Then _
                  rstAux.Close
               SQL = "select sum(valor_item) from LANCAMENTO WITH (NOLOCK)"
               SQL = SQL & " INNER JOIN ITEMLANCAMENTO WITH (NOLOCK)"
               SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
               SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
               SQL = SQL & " and status = 'A' "
               SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
               SQL = SQL & " and tipo_lancamento = 1"
               'SQL = SQL & " and formapagto_id <> 1"
               rstAux.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If Not rstAux.EOF Then _
                  If Not IsNull(rstAux.Fields(0).Value) Then _
                     VALOR_PENDENTE_N = 0 & rstAux.Fields(0).Value
               If rstAux.State = 1 Then _
                  rstAux.Close

               If VALOR_PENDENTE_N >= LIMITE_CREDITO_CLI_N Then
                  MsgBox "Valor limite de credito para esse cliente ultrapassado, não permitido venda, verificar com departamento financeiro."
                  CNPJ_CPF_A = ""
                  Exit Function
               End If
            End If
         End If

         If DiasAtrazoCliente_N > 0 Then
            VALOR_PENDENTE_N = 0
            DIAS_DE_ATRAZO = 0
            If rstAux.State = 1 Then _
               rstAux.Close
            SQL = "select min(DT_VENCIMENTO) from LANCAMENTO"
            SQL = SQL & " INNER JOIN ITEMLANCAMENTO"
            SQL = SQL & " ON LANCAMENTO.LANCAMENTO_ID = ITEMLANCAMENTO.LANCAMENTO_ID"
            SQL = SQL & " where pessoa_id = " & PESSOA_ID_N
            SQL = SQL & " and status = 'A'"
            SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
            SQL = SQL & " and tipo_lancamento = 1"
            rstAux.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not rstAux.EOF Then
               If Not IsNull(rstAux.Fields(0).Value) Then
                  DATA_INI = 0 & rstAux.Fields(0).Value
                  DATA_FIM = Date
                  DIAS_DE_ATRAZO = DATA_FIM - DATA_INI
               End If
            End If
            If rstAux.State = 1 Then _
               rstAux.Close

            If DIAS_DE_ATRAZO > 0 Then
               If DIAS_DE_ATRAZO > DiasAtrazoCliente_N Then
                  MsgBox "Cliente com parcelas em atrazo, verificar com ADM. Dias de atrazo = " & DIAS_DE_ATRAZO
                  Exit Function
               End If
            End If
         End If
      End If

      TRATA_PESSOA = True
      Else
         If INDR_PEDIDO_VENDA = False Then
            MsgBox "Cliente não encontrado !!!   "
            Else: TRATA_PESSOA = True
         End If
   End If   'fim select principal

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "TRATA_PESSOA"
End Function
'=====================
Public Function LE_PRODUTO(DADO_INFORMADO_CONSULTAR As String, TIPO_CONSULTA_A As String) As Boolean
'On Error GoTo ERRO_TRATA

'TIPO_CONSULTA_A = 'C' não mostrar mensagem de produto não cadastrador

   LE_PRODUTO = False
   DADO_INFORMADO_CONSULTAR = Trim(DADO_INFORMADO_CONSULTAR)
   PRODUTO_ID_N = 0
   QTDE_N = 0
   CRITERIO_A = ""
   INDR_PROD_BALANCA = False

   If Trim(DADO_INFORMADO_CONSULTAR) = "" Then _
      Exit Function

   If TabProduto.State = 1 Then _
      TabProduto.Close
   'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   SQL = "select count(produto_id) from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where codg_barra = '" & Trim(DADO_INFORMADO_CONSULTAR) & "'"
   'If TIPO_CONSULTA_A = "C" Then _
      SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      If Not IsNull(TabProduto.Fields(0).Value) Then
         If TabProduto.Fields(0).Value > 1 Then
            CRITERIO_A = Trim(DADO_INFORMADO_CONSULTAR)

            frmPEDIDOBARRAS.Show 1

            If Trim(CRITERIO_A) <> "" Then
               DADO_INFORMADO_CONSULTAR = "" & Trim(CRITERIO_A)
               LE_PRODUTO = True
               'MOSTRA_DADOS_PRODUTO
               'Exit Function
            End If
         End If
      End If
   End If

'LENDO PELO CODIGO DE BARRAS PRODUTOS REVENDA PRODUTOFORNECEDOR.codg_barra
'se tiver mais de um produto com o mesmo codigo de barras dai entra aqui para escolher qual produto vai vender
   INDR_LEU_POR_CODG_BARRAS = False
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from vwProduto WITH (NOLOCK)"
   SQL = SQL & " where barrafornec = '" & Trim(DADO_INFORMADO_CONSULTAR) & "'"
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      LE_PRODUTO = True
      INDR_LEU_POR_CODG_BARRAS = True
      MOSTRA_DADOS_PRODUTO
      Exit Function
   End If
   If TabProduto.State = 1 Then _
      TabProduto.Close

'LENDO PELO CODIGO DE BARRAS PRODUTOS REVENDA PRODUTO.codg_barra
   INDR_LEU_POR_CODG_BARRAS = False
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from vwProduto WITH (NOLOCK)"
   SQL = SQL & " where codg_barra = '" & Trim(DADO_INFORMADO_CONSULTAR) & "'"
   'If TIPO_CONSULTA_A = "C" Then _
      SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      LE_PRODUTO = True
      INDR_LEU_POR_CODG_BARRAS = True
      MOSTRA_DADOS_PRODUTO
      Exit Function
   End If

'LENDO PELO CODIGO DE BARRAS PRODUTOS DE PRODUÇÃO
   'le por codigo de barras ean 13 etiqueta balança
   If Len(DADO_INFORMADO_CONSULTAR) = 13 Then
      '2 = produtos "in store" (sempre será 2)     1
      'C = código do produto (4,5 ou 6 dígitos)    2 a 8
      'T = total a pagar (sempre 6 dígitos)        9 a 13
      'P = peso (sempre 5 dígitos)
      'Q = quantidade (sempre 5 dígitos)
      '0 = zero fixo
      'DV = dígito verificador do EAN-13

      'pegando codigo do produto no codigo de barras da etiqueta de balança
      CODIGO_BARRAS_A = "" & DADO_INFORMADO_CONSULTAR
      DADO_INFORMADO_CONSULTAR = "" & Int(Mid(DADO_INFORMADO_CONSULTAR, CasaInicioCodgProdBarra_N, TamanhoCodgProdBarra_N))

      If TabProduto.State = 1 Then _
         TabProduto.Close

      SQL = "select * from vwProduto WITH (NOLOCK)"
      SQL = SQL & " where codg_produto = '" & Trim(DADO_INFORMADO_CONSULTAR) & "'"
      'If TIPO_CONSULTA_A = "C" Then _
         SQL = SQL & " and situacao <> 'C' "
      TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabProduto.EOF Then
         If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then
            INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value
            QTDE_N = 1

            If Trim(DADO_INFORMADO_CONSULTAR) <> "" And INDR_PROD_BALANCA = True Then
               VALOR_ITEM_N = 0 & Mid(CODIGO_BARRAS_A, 8, TamanhoPesoValorBarra_N) / 100

               If UCase(PESO_VALOR_A) = UCase("GRAMAS") Then
                  QTDE_N = 0 & Int(Mid(CODIGO_BARRAS_A, 8, TamanhoPesoValorBarra_N))   'gramas
                  QTDE_N = QTDE_N / 1000
                  VALOR_ITEM_N = 0 & (TabProduto.Fields("PRECO_VENDA").Value * QTDE_N)

                  'regra: se o produto é de balança e unidade medida UN dai vai pegar unidade ao invez de peso
                  If Not IsNull(TabProduto.Fields("unidade_medida").Value) Then
                     If UCase(Trim(TabProduto.Fields("unidade_medida").Value)) = "UN" Then
                        QTDE_N = 0 & Int(Mid(CODIGO_BARRAS_A, 8, TamanhoPesoValorBarra_N))   'unidade
                        If Not IsNull(TabProduto.Fields("preco_venda").Value) Then _
                           VALOR_ITEM_N = TabProduto.Fields("preco_venda").Value
                     End If
                  End If
                  Else: QTDE_N = 0 & CONVERTE_VALOR_GRAMA(VALOR_ITEM_N, 0, TabProduto.Fields("produto_id").Value) 'sta
               End If
            End If

            LE_PRODUTO = True
            MOSTRA_DADOS_PRODUTO

            If TabProduto.State = 1 Then _
               TabProduto.Close

            Exit Function
            Else
               'If TIPO_CONSULTA_A = "C" Then _
                  MsgBox "Verificar cadastro produto."
         End If
         Else
            If TIPO_CONSULTA_A <> "C" Then _
               MsgBox "Verificar cadastro produto. " & DADO_INFORMADO_CONSULTAR
      End If   'TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   End If      'If Len(DADO_INFORMADO_CONSULTAR) = 13 Then
   If TabProduto.State = 1 Then _
      TabProduto.Close

'LENDO PELO CODIGO DO PRODUTO
   If TabProduto.State = 1 Then _
      TabProduto.Close

   SQL = "select * from vwProduto WITH (NOLOCK)"
   SQL = SQL & " where CODG_PRODUTO = '" & Trim(DADO_INFORMADO_CONSULTAR) & "'"
   If TIPO_CONSULTA_A <> "C" Then _
      SQL = SQL & " and situacao <> 'C' "
   TabProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabProduto.EOF Then
      LE_PRODUTO = True
      MOSTRA_DADOS_PRODUTO
      Exit Function
   End If

   If TIPO_CONSULTA_A <> "CADASTRO" Then
      MsgBox "Produto não cadastrado."
      Else
         LE_PRODUTO = True
         CODG_PRODUTO_A = "" & DADO_INFORMADO_CONSULTAR
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "LE_PRODUTO"
End Function

Public Sub MOSTRA_DADOS_PRODUTO()
'On Error GoTo ERRO_TRATA

   PRODUTO_ID_N = 0 & TabProduto.Fields("produto_id").Value

   INDR_PRODUTO_PRODUCAO_B = False 'verificando se o produto é de produção
   PESO_ITEM_N = QTDE_N

   INDR_PRODUTO_PRODUCAO_B = False
   If Not IsNull(TabProduto.Fields("producao").Value) Then _
      If TabProduto.Fields("producao").Value = True Then _
         INDR_PRODUTO_PRODUCAO_B = True

   INDR_PROD_BALANCA = False
   If Not IsNull(TabProduto.Fields("produto_balanca").Value) Then _
      INDR_PROD_BALANCA = TabProduto.Fields("produto_balanca").Value

   CODG_PRODUTO_A = "" & Trim(TabProduto.Fields("codg_produto").Value)
   DESC_PRODUTO_A = "" & Trim(TabProduto.Fields("descricao").Value)
   STATUS_PROD = "" & Trim(TabProduto.Fields("SITUACAO").Value)
   PR_CUSTO_PRODUTO_N = 0 & TabProduto.Fields("PRECO_CUSTO").Value
   PESO_LIQUIDO_N = 0 & TabProduto.Fields("peso_liquido").Value
   PR_ATACADO_N = 0 & TabProduto.Fields("PRECO_ATACADO").Value
   PR_VAREJO_N = 0 & TabProduto.Fields("PRECO_VENDA").Value
   CODG_NCM_A = "" & Trim(TabProduto.Fields("codg_ncm").Value)
   UNIDADE_MEDIDA_A = "" & Trim(TabProduto.Fields("unidade_medida").Value)
   ALIQUOTA_ICMS_N = 0 & Trim(TabProduto.Fields("ALIQUOTA_ICMS").Value)
   SITUACAO_TRIBUT_A = "" & Trim(TabProduto.Fields("SITUACAO_TRIBUTARIA").Value)
   REFERENCIA_A = "" & Trim(TabProduto.Fields("referencia").Value)

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "MOSTRA_DADOS_PRODUTO"
End Sub

Public Function FORMATA_CNPJCPF(CNPJCPF_A As String) As String
'On Error GoTo ERRO_TRATA

   FORMATA_CNPJCPF = ""
   If Trim(CNPJCPF_A) = "" Then _
      Exit Function

   If Len(Trim(CNPJCPF_A)) <= 11 Then
      FORMATA_CNPJCPF = "" & Left(CNPJCPF_A, 3) + "." + Mid(CNPJCPF_A, 4, 3) + "." + Mid(CNPJCPF_A, 7, 3) + "-" + Mid(CNPJCPF_A, 10, 2)
      Else: FORMATA_CNPJCPF = "" & Left(CNPJCPF_A, 2) + "." + Mid(CNPJCPF_A, 3, 3) + "." + Mid(CNPJCPF_A, 6, 3) + "." + Mid(CNPJCPF_A, 9, 4) + "." + Mid(CNPJCPF_A, 13, 2)
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "FORMATA_CNPJCPF"
End Function

Public Sub GRAVA_NOTA(Numr_Nota_N As String, _
                      SERIE_DOC_A As String, _
                      Modelo_Doc_A As String, _
                      NF_TIPO_A As String, _
                      QTDE_RODAPE_A As String, _
                      PESO_BRUTO_A As String, _
                      PESO_LIQUI_A As String, _
                      INDPRES_N As String, _
                      INDDEST_N As String, _
                      CFOP_AUX_A As String, _
                      CNPJCPF_TRANS_A As String)
'On Error GoTo ERRO_TRATA

   Dim TRANSP_ID_N  As Long

   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   SQL = "select NF.nf_id from PEDIDO WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PEDIDONF WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDO.PEDIDO_ID = PEDIDONF.PEDIDO_ID"
   SQL = SQL & " INNER JOIN NF WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDONF.NF_ID = NF.NF_ID "

   SQL = SQL & " where PEDIDO.pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and NF.estabelecimento_id = " & ESTABELECIMENTO_ID_N
   SQL = SQL & " and modelo_doc = '" & Trim(Modelo_Doc_A) & "'"

   TabNOTA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabNOTA.EOF Then
      NF_ID_N = TabNOTA.Fields("nf_id").Value
      If TabNOTA.State = 1 Then _
         TabNOTA.Close

      SQL = "delete from NFITEM "
      SQL = SQL & " where nf_id = " & NF_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "delete from PEDIDONF "
      SQL = SQL & " where nf_id = " & NF_ID_N
      CONECTA_RETAGUARDA.Execute SQL

      SQL = "delete from NF "
      SQL = SQL & " where nf_id = " & NF_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabNOTA.State = 1 Then _
      TabNOTA.Close

   If Trim(CNPJCPF_TRANS_A) = "" Then _
      CNPJCPF_TRANS_A = "" & CNPJ_EMPRESA_N

   If Trim(QTDE_RODAPE_A) = "" Then _
      QTDE_RODAPE_A = "1"

   If Trim(ESPECIE_A) = "" Then _
      ESPECIE_A = "UN"

   If Trim(PESO_BRUTO_A) = "" Then _
      PESO_BRUTO_A = "0"

   If Trim(PESO_LIQUI_A) = "" Then _
      PESO_LIQUI_A = "0"

   TRANSP_ID_N = 0 & TRAZ_ID_TABELA("vwTRANSPORTADORA", "transp_id", "cnpjcpf", CNPJCPF_TRANS_A)

   SQL3 = ESTABELECIMENTO_ID_N

   NF_ID_N = 1
   'nf_id_n = MAX_ID("nf_id", "nf", "estabelecimento_id", SQL3, "", "")
   NF_ID_N = MAX_ID("nf_id", "nf", "", "", "", "")

'rever aqui se vai funcionar, acho que tem que travar a tabela pra não gerar o mesmo id pra vendas diferentes
   SQL = "INSERT INTO NF ("
      SQL = SQL & " NF_ID, NF_TIPO, NUMR_NOTA, SERIE_NOTA, "
      SQL = SQL & " estabelecimento_id, DT_EMISSAO, DT_ENTRASAI, TRANSP_ID, "
      SQL = SQL & " Qtd_Volume, Peso_Bruto, Peso_Liquido, status,pessoa_id,"
      SQL = SQL & " indPres,idDest,modelo_doc"
   SQL = SQL & " )"
   SQL = SQL & " VALUES ("
      SQL = SQL & NF_ID_N                           'NF_ID
      SQL = SQL & ",'" & Trim(NF_TIPO_A) & "'"              'NF_TIPO
      SQL = SQL & "," & Numr_Nota_N                         'NUMR_NOTA
      SQL = SQL & ",'" & Trim(SERIE_DOC_A) & "'"            'SERIE_NOTA
      SQL = SQL & "," & ESTABELECIMENTO_ID_N                'estabelecimento_id
      SQL = SQL & ",'" & Now & "'"                          'DT_EMISSAO
      SQL = SQL & ",'" & Now & "'"                          'DT_ENTRASAI
      SQL = SQL & "," & TRANSP_ID_N                         'TRANSP_ID
      SQL = SQL & "," & Replace(QTDE_RODAPE_A, ",", ".")    'Qtd_Volume
      SQL = SQL & "," & Replace(PESO_BRUTO_A, ",", ".")     'Peso_Bruto
      SQL = SQL & "," & Replace(PESO_LIQUI_A, ",", ".")     'Peso_Liquido
      SQL = SQL & ",'" & "i" & "'"                          'status
      SQL = SQL & "," & PESSOA_ID_N                         'pessoa_id
      SQL = SQL & "," & INDPRES_N                           'indPres
      SQL = SQL & "," & INDDEST_N                           'idDest
      SQL = SQL & ",'" & Trim(Modelo_Doc_A) & "'"           'modelo_doc
   SQL = SQL & " )"
   CONECTA_RETAGUARDA.Execute SQL

   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

   SQL = "SELECT PEDIDOITEM.*, PRODUTO.CODG_PRODUTO, PRODUTO.DESCRICAO, PRODUTO.SITUACAO"
   SQL = SQL & " FROM PEDIDOITEM  WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN PRODUTO  WITH (NOLOCK)"
   SQL = SQL & " ON PEDIDOITEM.PRODUTO_ID = PRODUTO.PRODUTO_ID"
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   'SQL = SQL & " and tipo_reg = 'PC' "
   SQL = SQL & " and pedidoitem.status <> 'C' "
   TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabPedidoItem.EOF Then
      PERC_DESCONTO_N = 0 & TabPedidoItem!PERC_DESC
      'TabPedidoItem.MoveFirst
      While Not TabPedidoItem.EOF
      If Trim(TabPedidoItem.Fields("situacao").Value) = "A" Then
         If TabAUX.State = 1 Then _
            TabAUX.Close

         If Not IsNull(TabPedidoItem!CFOP_ID) Then _
            If Trim(TabPedidoItem!CFOP_ID) <> "" Then _
               If IsNumeric(TabPedidoItem!CFOP_ID) Then _
                  CFOP_ID_N = Trim(TabPedidoItem!CFOP_ID)

         'diversas
         If Trim(NF_TIPO_A) = "DV" Then _
             CFOP_ID_N = Trim(CFOP_AUX_A)

         'cupom fiscal
         If Trim(CFOP_AUX_A) = 5929 Or Trim(CFOP_AUX_A) = 6929 Then _
            CFOP_ID_N = Trim(CFOP_AUX_A)

         If CFOP_ID_N <= 0 Then
            CFOP_ID_N = CFOP_AUX_A
            Else: CFOP_AUX_A = "" & CFOP_ID_N
         End If

         SQL = "select * from NFITEM WITH (NOLOCK)"
         SQL = SQL & " where nf_id = " & NF_ID_N
         SQL = SQL & " and produto_id = " & TabPedidoItem.Fields("produto_id").Value
         SQL = SQL & " and seq_id = " & TabPedidoItem.Fields("seq_id").Value
         TabAUX.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabAUX.EOF Then
            SQL = "UPDATE NFITEM SET "
               SQL = SQL & " Valor = " & tpMOEDA(TabPedidoItem!Valor_Item - (TabPedidoItem!Valor_Item * PERC_DESCONTO_N / 100))
               SQL = SQL & ", Desconto = " & tpMOEDA((TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA) * PERC_DESCONTO_N / 100)
               SQL = SQL & ", Qtde = " & tpMOEDA(TabPedidoItem!QTD_PEDIDA)
               SQL = SQL & ", CFOP_id = '" & CFOP_ID_N & "'"
               SQL = SQL & ", STRIBUTARIA = " & tpMOEDA(TabPedidoItem!STRIBUTARIA)
               SQL = SQL & ", VlrBaseIcms = " & tpMOEDA(TabPedidoItem!VLRBASEICMS)
               SQL = SQL & ", PERCICMS = " & tpMOEDA(TabPedidoItem!PERCICMS)
               SQL = SQL & ", VlrICMS = " & tpMOEDA(TabPedidoItem!VLRICMS)
               SQL = SQL & ", VLRBASEICMSSUBST = " & tpMOEDA(TabPedidoItem!VLRBASEICMSSUBST)
               SQL = SQL & ", PERCICMSSUBST = " & tpMOEDA(TabPedidoItem!PERCICMSSUBST)
               SQL = SQL & ", VLRICMSSUBST = " & tpMOEDA(TabPedidoItem!VLRICMSSUBST)
               SQL = SQL & ", PERCREDUCAOICMS = " & tpMOEDA(TabPedidoItem!PERCREDUCAOICMS)
               SQL = SQL & ", PERCIVA = " & tpMOEDA(TabPedidoItem!PERCIVA)
            SQL = SQL & " where nf_id = " & NF_ID_N
            SQL = SQL & " and produto_id = " & TabPedidoItem.Fields("produto_id").Value
            SQL = SQL & " and seq_id = " & TabPedidoItem.Fields("seq_id").Value
            Else
               SQL = "INSERT INTO NFITEM ("
                  SQL = SQL & "nf_id, seq_id, produto_id, Valor, Desconto, Qtde, CFOP_id, STRIBUTARIA, "
                  SQL = SQL & "VlrBaseIcms, PERCICMS, VlrICMS,  VLRBASEICMSSUBST, PERCICMSSUBST, "
                  SQL = SQL & "VLRICMSSUBST, PERCREDUCAOICMS, PERCIVA, PERC_IPI"
               SQL = SQL & ")"
               SQL = SQL & " VALUES ("
                  SQL = SQL & NF_ID_N                                                                                             'nf_id
                  SQL = SQL & "," & TabPedidoItem.Fields("seq_id").Value
                  SQL = SQL & "," & TabPedidoItem.Fields("produto_id").Value
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!Valor_Item - (TabPedidoItem!Valor_Item * PERC_DESCONTO_N / 100))  'Valor
                  SQL = SQL & "," & tpMOEDA((TabPedidoItem!Valor_Item * TabPedidoItem!QTD_PEDIDA) * PERC_DESCONTO_N / 100)  'Desconto
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!QTD_PEDIDA)                                                               'Qtde
                  SQL = SQL & ",'" & CFOP_ID_N & "'"                                                                                'CFOP_id
                  SQL = SQL & ",'" & TabPedidoItem!STRIBUTARIA & "'"                                                                'STRIBUTARIA
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!VLRBASEICMS)                                                              'VlrBaseIcms
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!PERCICMS)                                                                 'PERCICMS
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!VLRICMS)                                                                  'VlrICMS
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!VLRBASEICMSSUBST)                                                         'VLRBASEICMSSUBST
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!PERCICMSSUBST)                                                            'PERCICMSSUBST
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!VLRICMSSUBST)                                                             'VLRICMSSUBST
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!PERCREDUCAOICMS)                                                          'PERCREDUCAOICMS
                  SQL = SQL & "," & tpMOEDA(TabPedidoItem!PERCIVA)                                                                  'PERCIVA
                  SQL = SQL & "," & 0                                                                                               'PERC_IPI
               SQL = SQL & ")"
         End If
         If TabAUX.State = 1 Then _
            TabAUX.Close

         CONECTA_RETAGUARDA.Execute SQL
      End If
         TabPedidoItem.MoveNext
      Wend

      GRAVA_STATUS_EMITIDO

      If TabPedidoItem.State = 1 Then _
         TabPedidoItem.Close

'=PEDIDONF
      
      SQL = "select * from PEDIDONF "
      SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
      TabPedidoItem.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If TabPedidoItem.EOF Then

         SQL = "insert into PEDIDONF "
            SQL = SQL & " (PEDIDONF_ID,PEDIDO_ID,NF_ID)"
         SQL = SQL & " values("
            SQL = SQL & MAX_ID("PEDIDONF_id", "PEDIDONF", "", "", "", "")
            SQL = SQL & "," & PEDIDO_ID_N
            SQL = SQL & "," & NF_ID_N
         SQL = SQL & " )"

         CONECTA_RETAGUARDA.Execute SQL
   End If
'=

   End If
   If TabPedidoItem.State = 1 Then _
      TabPedidoItem.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "GRAVA_NOTA"
End Sub

Public Sub GRAVA_STATUS_EMITIDO()
'On Error GoTo ERRO_TRATA

   SQL = "update PEDIDO set "
   SQL = SQL & "status = 3 " 'foi Gerado Nota Fiscal
   SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "GRAVA_STATUS_EMITIDO"
End Sub

Public Function GERA_NUMERO_NFe_N() As Long
'On Error GoTo ERRO_TRATA

   Dim EMPRESA_ID_A  As String
   Dim ESTAB_ID_A    As String
   Dim stsSQL        As String

   EMPRESA_ID_A = EMPRESA_ID_N
   ESTAB_ID_A = ESTABELECIMENTO_ID_N
   'GERA_NUMERO_NFe_N = 0 & MAX_ID("seq_nota_saida", "ESTABELECIMENTO", "empresa_id", EMPRESA_ID_A, "", ")
   'stsSQL = "update EMPRESA set "
   'stsSQL = stsSQL & " seq_nota_saida = " & GERA_NUMERO_NFe_N
   'stsSQL = stsSQL & " from EMPRESA "
   'stsSQL = stsSQL & " INNER JOIN ESTABELECIMENTO "
   'stsSQL = stsSQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   'stsSQL = stsSQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   'stsSQL = stsSQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   'CONECTA_RETAGUARDA.Execute stsSQL

   GERA_NUMERO_NFe_N = 0 & MAX_ID("seq_nota_saida", "ESTABELECIMENTO", "empresa_id", EMPRESA_ID_A, "ESTABELECIMENTO_ID", ESTAB_ID_A)

   stsSQL = "update ESTABELECIMENTO set "
   stsSQL = stsSQL & " seq_nota_saida = " & GERA_NUMERO_NFe_N
   stsSQL = stsSQL & " where empresa_id = " & EMPRESA_ID_N
   stsSQL = stsSQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute stsSQL

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "GERA_NUMERO_NFe_N"
End Function

Public Function GERA_NUMERO_NFC_N(MODELO_DOC As String) As Long
'On Error GoTo ERRO_TRATA

   Dim EMPRESA_ID_A  As String
   Dim stsSQL        As String

   EMPRESA_ID_A = EMPRESA_ID_N

'VERIFICANDO SE JA TEM CUPOM GERADO NO GLOBAL
VERIFICA_TRAVEIS:
   GERA_NUMERO_NFC_N = 0 & MAX_ID("seq_cupom", "ESTABELECIMENTO", "estabelecimento_id", Str(ESTABELECIMENTO_ID_N), "", "")

   stsSQL = "update ESTABELECIMENTO set "
   stsSQL = stsSQL & " seq_cupom = " & GERA_NUMERO_NFC_N
   stsSQL = stsSQL & " from ESTABELECIMENTO "
   stsSQL = stsSQL & " where estabelecimento_id = " & ESTABELECIMENTO_ID_N
   CONECTA_RETAGUARDA.Execute stsSQL

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then
      MsgBox "Banco GLOBAL não conectado."
      Exit Function
      Else
         If TabTemp.State = 1 Then _
            TabTemp.Close
         SQL = "select MFASEQUENCIA from MFA010 WITH (NOLOCK) "
         SQL = SQL & " where mfadoc = '" & Trim(GERA_NUMERO_NFC_N) & "'"

         SQL = SQL & " and MFALOJA = '0" & ESTABELECIMENTO_ID_N & "'"
         SQL = SQL & " and MFAfilial = '0" & ESTABELECIMENTO_ID_N & "'"

         SQL = SQL & " and mfaprefixo = 'NFC'"
         TabTemp.Open SQL, CONECTA_GLOBAL, , , adCmdText
         If Not TabTemp.EOF Then _
            If Not IsNull(TabTemp.Fields(0).Value) Then _
               GoTo VERIFICA_TRAVEIS
         If TabTemp.State = 1 Then _
            TabTemp.Close
   End If

   GRAVA_CUPOM GERA_NUMERO_NFC_N, MODELO_DOC

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "GERA_NUMERO_NFC_N"
End Function

Public Function GRAVA_CUPOM(Numero_Cupom_N As Long, _
                            Modelo_Doc_A As String) As Long
'On Error GoTo ERRO_TRATA

   GRAVA_CUPOM = 0

   Dim TabCupom   As New ADODB.Recordset
   Dim strSQL     As String

   If IMPRESSORA_ID_N <= 0 Then _
      IMPRESSORA_ID_N = 1

   If TabCupom.State = 1 Then _
      TabCupom.Close

   'GRAVA TABELA CUPOM
   strSQL = "select * from CUPOM WITH (NOLOCK) "
   strSQL = strSQL & " where numr_cupom = " & Numero_Cupom_N
   strSQL = strSQL & " and modelo_doc = '" & Trim(Modelo_Doc_A) & "'"
   strSQL = strSQL & " and pedido_id = " & PEDIDO_ID_N
   TabCupom.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
   If TabCupom.EOF Then
      If TabCupom.State = 1 Then _
         TabCupom.Close

      GRAVA_CUPOM = 0 & MAX_ID("cupom_id", "cupom", "", "", "", "")

      strSQL = "insert into CUPOM "
      strSQL = strSQL & " (CUPOM_ID,PEDIDO_ID,NUMR_CUPOM,MODELO_DOC)"
      strSQL = strSQL & " VALUES("
         strSQL = strSQL & GRAVA_CUPOM                'CUPOM_ID
         strSQL = strSQL & "," & PEDIDO_ID_N          'PEDIDO_ID
         strSQL = strSQL & "," & Numero_Cupom_N       'NUMR_CUPOM
         strSQL = strSQL & ",'" & Modelo_Doc_A & "'"  'MODELO_DOC
      strSQL = strSQL & ")"
      CONECTA_RETAGUARDA.Execute strSQL
   End If
   If TabCupom.State = 1 Then _
      TabCupom.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "GRAVA_CUPOM"
End Function

Public Function TRAZ_ST_PRODUTO(PRODUTO_ID As Long, CODG_PROD As String) As String
'On Error GoTo ERRO_TRATA

   TRAZ_ST_PRODUTO = ""

   If PRODUTO_ID <= 0 And Trim(CODG_PROD) = "" Then _
      Exit Function

   If Pessoa_id > 0 Then
      Dim TabRG   As New ADODB.Recordset
      Dim strSQL  As String

      If TabRG.State = 1 Then _
         TabRG.Close

      strSQL = "select NUMERO_RG from RG WITH (NOLOCK)"
      strSQL = strSQL & " where pessoa_id = " & Pessoa_id
      TabRG.Open strSQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabRG.EOF Then _
         TRAZ_ST_PRODUTO = "" & Trim(TabRG.Fields("NUMERO_RG").Value)
      If TabRG.State = 1 Then _
         TabRG.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "TRAZ_ST_PRODUTO"
End Function

Public Sub CHECA_CLIENTE(CNPJCPF_A As String, NOME_A As String)
'On Error GoTo ERRO_TRATA

   If Trim(CNPJCPF_A) <> "" And Trim(NOME_A) <> "" Then
      CLIENTE_ID_N = 0
      PESSOA_ID_N = 0

      If TabCliente.State = 1 Then _
         TabCliente.Close

      SQL = "select cliente_id,pessoa_id,status from CLIENTE WITH (NOLOCK)"
      SQL = SQL & " where cgccpf = '" & Trim(CNPJCPF_A) & "'"
      TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCliente.EOF Then
         CLIENTE_ID_N = 0 & TabCliente.Fields("cliente_id").Value
         PESSOA_ID_N = 0 & TabCliente.Fields("pessoa_id").Value
         If IsNull(TabCliente.Fields("status").Value) Then
            MsgBox "Problemas no cadastro do cliente."
            Else
               If Trim(UCase(TabCliente.Fields("status").Value)) <> "A" Then
                  MsgBox "Problemas no cadastro do cliente."
               End If
         End If
         Else  'só vai inserir
            If TabCliente.State = 1 Then _
               TabCliente.Close

            SQL = "select pessoa_id from PESSOA WITH (NOLOCK)"
            SQL = SQL & " where cnpjcpf = '" & Trim(CNPJCPF_A) & "'"
            TabCliente.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If Not TabCliente.EOF Then
               PESSOA_ID_N = 0 & TabCliente.Fields("pessoa_id").Value
               Else: spPessoa 1, 0, Trim(CNPJCPF_A), Trim(NOME_A), Trim(NOME_A), "A"
            End If
            If TabCliente.State = 1 Then _
               TabCliente.Close

            PESSOA_ID_N = 0 & TRAZ_ID_TABELA("PESSOA", "PESSOA_ID", "CNPJCPF", Trim(CNPJCPF_A))
            CLIENTE_ID_N = MAX_ID("cliente_id", "cliente", "", "", "", "")

            spCliente 1, 1, Trim(CNPJCPF_A), Trim(NOME_A), Trim(NOME_A), "", Now, "A", "", "", "", "", "", "0", "", "0", "", ""
      End If
      If TabCliente.State = 1 Then _
         TabCliente.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "CHECA_CLIENTE"
End Sub

Public Sub spEncomenda(Acao_N As Long, PEDIDOENCOMENDA_ID_N As Long, PEDIDO_ID As Long, USUARIO_ID As Long, VLR_TX_ENTREGA As Double)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   If ENCOMENDA_ID_N <= 0 Then _
      ENCOMENDA_ID_N = MAX_ID("PEDIDOENCOMENDA_ID", "PEDIDOENCOMENDA", "", "", "", "")

   SQL = "spPEDIDOENCOMENDA " & Acao_N & "," & ENCOMENDA_ID_N & "," & PEDIDO_ID & "," & USUARIO_ID & "," & tpMOEDA(VLR_TX_ENTREGA)

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spEncomenda"
End Sub

Public Sub EXCLUIR_400()
On Error Resume Next

   If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

   ABRE_BANCO_GLOBAL

   If CONECTA_GLOBAL.State <> 1 Then _
      Exit Sub

   Dim TabDir     As New ADODB.Recordset
   Dim Driver_A   As String
   Dim Driver_B   As String

   Driver_A = ""
   If TabDir.State = 1 Then _
      TabDir.Close

   SQL = "select envionfe from EMPRES"
   TabDir.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabDir.EOF Then
      If Not IsNull(TabDir.Fields(0).Value) Then
         Driver_A = "" & TabDir.Fields(0).Value
         Driver_B = "" & TabDir.Fields(0).Value
      End If
   End If
   If TabDir.State = 1 Then _
      TabDir.Close

   Driver_A = Left(Driver_A, 1) & ":\NFE\nfe\wsdl\Homologacao\GO"
   Driver_B = Left(Driver_B, 1) & ":\NFE\nfe\wsdl\producao\GO"

   Dim Arquivo       As File
   Dim SubDiretorio  As Folder
   Dim DiretoriO     As Folder

   Set DiretoriO = FSO.GetFolder(Driver_A)

   For Each Arquivo In DiretoriO.Files
      If Trim(Right(FSO.GetBaseName(Arquivo.Name), 4)) = "_400" Then
         FSO.DeleteFile (Arquivo)
      End If
   Next

   Set DiretoriO = FSO.GetFolder(Driver_B)

   For Each Arquivo In DiretoriO.Files
      If Trim(Right(FSO.GetBaseName(Arquivo.Name), 4)) = "_400" Then
         FSO.DeleteFile (Arquivo)
      End If
   Next

End Sub

Public Function MOSTRA_VERSAO_NFe(CNPJ_A As String) As String
'On Error GoTo ERRO_TRATA

   MOSTRA_VERSAO_NFe = ""

   If CONECTA_GLOBAL.State <> 1 Then
      'CONECTA_GLOBAL.Close
      ABRE_BANCO_GLOBAL
   End If

   If CONECTA_GLOBAL.State <> 1 Then
      'MsgBox "Banco GLOBAL não conectado."
      Exit Function
   End If

   Dim TabEmpres As New ADODB.Recordset

   If TabEmpres.State = 1 Then _
      TabEmpres.Close

   SQL = "select versaonfe from EMPRES "
   SQL = SQL & " where cnpj = '" & Trim(CNPJ_A) & "'"
   TabEmpres.Open SQL, CONECTA_GLOBAL, , , adCmdText
   If Not TabEmpres.EOF Then _
      If Not IsNull(TabEmpres.Fields(0).Value) Then _
         MOSTRA_VERSAO_NFe = "" & TabEmpres.Fields(0).Value
   If TabEmpres.State = 1 Then _
      TabEmpres.Close
   'If CONECTA_GLOBAL.State = 1 Then _
      CONECTA_GLOBAL.Close

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "MOSTRA_VERSAO_NFe"
End Function

Public Function CHECAR_CAIXA(USU_ID_N As Long, DT_ABERTURA As String) As Boolean
'On Error GoTo ERRO_TRATA

   CHECAR_CAIXA = False

   If Trim(USU_ID_N) <> "" And Trim(DT_ABERTURA) <> "" Then
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close

      SQL = "SELECT CAIXADIA_ID,DT_ABERTURA,DT_FECHAMENTO FROM CAIXADIA WITH (NOLOCK)"

      SQL = SQL & " where USUARIO_ID = " & USU_ID_N

      SQL = SQL & " and dt_abertura >= '" & DMA(DT_ABERTURA, "i") & "'"
      SQL = SQL & " and dt_abertura <= '" & DMA(DT_ABERTURA, "f") & "'"

      TabCAIXA.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabCAIXA.EOF Then
         If Not IsNull(TabCAIXA.Fields("DT_FECHAMENTO").Value) Then
            If IsDate(TabCAIXA.Fields("DT_FECHAMENTO").Value) Then
               If CDate(TabCAIXA.Fields("DT_FECHAMENTO").Value) >= Date Then
                  MsgBox "Caixa já fechado. Data Fechamento : " & TabCAIXA.Fields("DT_FECHAMENTO").Value
                  Exit Function
                  If TabCAIXA.State = 1 Then _
                     TabCAIXA.Close
               End If
            End If
         End If
         CHECAR_CAIXA = True
         Else  'NÃO ABRIU CAIXA PARA USUARIO
            CHECAR_CAIXA = False

            If TabCAIXA.State = 1 Then _
               TabCAIXA.Close

            If frmCaixa.Visible = False Then
               Msg = "Caixa não aberto para este usuário, Deseja realizar abertura agora?"
               PERGUNTA Msg, vbYesNo + 32, "Caixa Balcão", "DEMO.HLP", 1000
               Msg = ""
               If RESPOSTA = vbYes Then _
                  frmCaixa.Show 1
            End If
      End If
      If TabCAIXA.State = 1 Then _
         TabCAIXA.Close
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "CHECAR_CAIXA"
End Function

Public Function CONSULTA_CEP_WEB(CEP_A As String) As Boolean
'On Error GoTo ERRO_TRATA
On Error Resume Next

   If Trim(CEP_A) = "" Then _
      Exit Function

   Dim myXML   As DOMDocument
   Dim X       As IXMLDOMNode
   Dim xlink   As String

   CONSULTA_CEP_WEB = False
   Xrua_A = ""
   Xuf_A = ""
   Xcidade_A = ""
   Xbairro_A = ""
   Xtipo_A = ""
   TxtLogradouro_A = ""

   MousePointer = 11
   Set myXML = New DOMDocument
   myXML.resolveExternals = True
   myXML.validateOnParse = True
   myXML.async = False
   xlink = "http://cep.republicavirtual.com.br/web_cep.php?cep=" & CEP_A & "&formato=xml"
   myXML.Load (xlink)
   For Each X In myXML.documentElement.childNodes
      Select Case X.nodeName
         Dim Xrua    As String
         Dim xtipo   As String
         Case Is = "logradouro"
              Xrua_A = UCase$(X.childNodes(0).Text)
         Case Is = "uf"
              Xuf_A = "" & UCase$(X.childNodes(0).Text)
         Case Is = "cidade"
              Xcidade_A = "" & UCase$(X.childNodes(0).Text)
         Case Is = "bairro"
              Xbairro_A = "" & UCase$(X.childNodes(0).Text)
         Case Is = "tipo_logradouro"
              Xtipo_A = "" & UCase$(X.childNodes(0).Text)
         Case Is = "resultado_txt"
              xresultado_A = "" & UCase$(X.childNodes(0).Text)
      End Select
      TxtLogradouro_A = Xtipo_A & " " & Xrua_A
   Next
   MousePointer = 0
   If Trim(Xuf_A) <> "" Then _
      CONSULTA_CEP_WEB = True

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "CONSULTA_CEP_WEB"
End Function
'==============================
Public Sub GERA_PEDIDO_ID()
'On Error GoTo ERRO_TRATA

   Dim TabGeraPedido  As New ADODB.Recordset

   If INDR_SEQUENCIA = True Then 'AQUI SEGUE ESQUEMA DE REAPROVEIDAR SEQUENCIA DE PEDIDO_ID
      PEDIDO_ID_N = 1
      CONT_N = 999999999
   
      For i = 1 To CONT_N '> PEDIDO_ID_N
         PEDIDO_ID_N = i
         If TabGeraPedido.State = 1 Then _
            TabGeraPedido.Close
   
         SQL = "select pedido_id from PEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabGeraPedido.EOF Then
            If TabGeraPedido.State = 1 Then _
               TabGeraPedido.Close
   
            SQL = "select numr_doc from LANCAMENTO WITH (NOLOCK)"
            SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
            TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabGeraPedido.EOF Then
               If TabGeraPedido.State = 1 Then _
                  TabGeraPedido.Close
   
               SQL = "select pedido_id from PEDIDOTEMP WITH (NOLOCK)"
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabGeraPedido.EOF Then
                  If TabGeraPedido.State = 1 Then _
                     TabGeraPedido.Close
   
                  SQL = "select os_id from OS WITH (NOLOCK)"
                  SQL = SQL & " where os_id = " & PEDIDO_ID_N
                  TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabGeraPedido.EOF Then
                     If TabGeraPedido.State = 1 Then _
                        TabGeraPedido.Close

                     SQL = "select entrada_id from NOTAENTRADA WITH (NOLOCK)"
                     SQL = SQL & " where entrada_id = " & PEDIDO_ID_N
                     TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If TabGeraPedido.EOF Then
                        If TabGeraPedido.State = 1 Then _
                           TabGeraPedido.Close

                        SQL = "select pedido_id from PEDIDOITEM WITH (NOLOCK)"
                        SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                        TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                        If TabGeraPedido.EOF Then
                           If TabGeraPedido.State = 1 Then _
                              TabGeraPedido.Close

                           SQL = "select pedido_id from CUPOM WITH (NOLOCK)"
                           SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
                           TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                           If TabGeraPedido.EOF Then _
                              Exit For
                        End If
                     End If
                  End If
               End If
            End If
         End If
         If TabGeraPedido.State = 1 Then _
            TabGeraPedido.Close
      Next

'MsgBox "pedido gerado = " & PEDIDO_ID_N

      SQL = "update EMPRESAPARAMETRO set "
      SQL = SQL & " seq_pedido = " & PEDIDO_ID_N
      CONECTA_RETAGUARDA.Execute SQL
   
      If TabGeraPedido.State = 1 Then _
         TabGeraPedido.Close
   
      SQL = "select seq_pedido from EMPRESAPARAMETRO WITH (NOLOCK)"
      TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
      If Not TabGeraPedido.EOF Then _
         If Not IsNull(TabGeraPedido.Fields(0).Value) Then _
            PEDIDO_ID_N = TabGeraPedido.Fields(0).Value
      If TabGeraPedido.State = 1 Then _
         TabGeraPedido.Close
      Else  'SEGUE SEQUENCIAL DA TABELA EMPRESA SEQ_PEDIDO
         PEDIDO_ID_N = 1

            SQL = "update EMPRESA set "
            SQL = SQL & " seq_pedido = seq_pedido + 1 "
            SQL = SQL & " where empresa.empresa_id = " & 1
            CONECTA_RETAGUARDA.Execute SQL

         If TabGeraPedido.State = 1 Then _
            TabGeraPedido.Close
         SQL = "select seq_pedido from EMPRESA WITH (NOLOCK)"
         SQL = SQL & " where empresa.empresa_id = " & 1
         TabGeraPedido.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If Not TabGeraPedido.EOF Then _
            If Not IsNull(TabGeraPedido.Fields(0).Value) Then _
               PEDIDO_ID_N = TabGeraPedido.Fields(0).Value
         If TabGeraPedido.State = 1 Then _
            TabGeraPedido.Close
   End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_PEDIDO_ID"
End Sub

Public Sub GERA_PEDIDO_ID_old()
'On Error GoTo ERRO_TRATA

   Dim TabEmp  As New ADODB.Recordset
   PEDIDO_ID_N = 1

'set
'aqui tem que parametrizar


   If Trim(CNPJ_EMPRESA_N) = "15333554000188" Then
      PEDIDO_ID_N = 1
      CONT_N = 999999999

      For i = 1 To CONT_N '> PEDIDO_ID_N
         PEDIDO_ID_N = i
         If TabEmp.State = 1 Then _
            TabEmp.Close

         SQL = "select pedido_id from PEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabEmp.EOF Then
            If TabEmp.State = 1 Then _
               TabEmp.Close

            SQL = "select numr_doc from LANCAMENTO WITH (NOLOCK)"
            SQL = SQL & " where numr_doc = " & PEDIDO_ID_N
            TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabEmp.EOF Then
               If TabEmp.State = 1 Then _
                  TabEmp.Close

               SQL = "select pedido_id from PEDIDOTEMP WITH (NOLOCK)"
               SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
               TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
               If TabEmp.EOF Then
                  If TabEmp.State = 1 Then _
                     TabEmp.Close

                  SQL = "select os_id from OS WITH (NOLOCK)"
                  SQL = SQL & " where os_id = " & PEDIDO_ID_N
                  TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                  If TabEmp.EOF Then
                     If TabEmp.State = 1 Then _
                        TabEmp.Close

                     SQL = "select entrada_id from NOTAENTRADA WITH (NOLOCK)"
                     SQL = SQL & " where entrada_id = " & PEDIDO_ID_N
                     TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
                     If TabEmp.EOF Then _
                        Exit For
                  End If
               End If
            End If
         End If
         If TabEmp.State = 1 Then _
            TabEmp.Close

         SQL = "select pedido_id from PEDIDO WITH (NOLOCK)"
         SQL = SQL & " where pedido_id = " & PEDIDO_ID_N
         TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
         If TabEmp.EOF Then _
            Exit For
      Next

      SQL = "update EMPRESA set "
      SQL = SQL & " seq_pedido = " & PEDIDO_ID_N
      SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
      CONECTA_RETAGUARDA.Execute SQL
      Else
         SQL = "update EMPRESA set "
         SQL = SQL & " seq_pedido = seq_pedido + 1 "
         SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
         CONECTA_RETAGUARDA.Execute SQL
   End If
   If TabEmp.State = 1 Then _
      TabEmp.Close

   'SQL = "select seq_pedido from EMPRESA WITH (LOCK)"
   'SQL = SQL & " INNER JOIN ESTABELECIMENTO WITH (LOCK)"
   'SQL = SQL & " ON EMPRESA.EMPRESA_ID = ESTABELECIMENTO.EMPRESA_ID"
   'SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   'SQL = SQL & " and estabelecimento_id = " & ESTABELECIMENTO_ID_N
   'TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   'If Not TabEmp.EOF Then _
      If Not IsNull(TabEmp.Fields(0).Value) Then _
         PEDIDO_ID_N = TabEmp.Fields(0).Value
   'If TabEmp.State = 1 Then _
      TabEmp.Close

   SQL = "select seq_pedido from EMPRESA "
   SQL = SQL & " where empresa.empresa_id = " & EMPRESA_ID_N
   TabEmp.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not TabEmp.EOF Then _
      If Not IsNull(TabEmp.Fields(0).Value) Then _
         PEDIDO_ID_N = TabEmp.Fields(0).Value
   If TabEmp.State = 1 Then _
      TabEmp.Close

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "GERA_PEDIDO_ID"
End Sub
'===================================
Public Sub spPedido(Acao_N As Integer, CGCCPF As String, DT_REQ As Date, STATUS As Integer, _
                    TIPO_REGISTRO As String, USUARIO_LIBERA_VENDA As Integer, Valor_Desconto As Double, _
                    PERC_DESC As Double, NOME_CLIENTE As String, VALOR_RECEBIDO As Double, _
                    VALOR_TOTAL As Double, PREFIXO As String, TABELAPRECO_ID_N As Integer)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   SQL = "spPedido " & Acao_N & "," & _
                       PEDIDO_ID_N & "," & CLIENTE_ID_N & "," & EMPRESA_ID_N & "," & _
                       ESTABELECIMENTO_ID_N & "," & CARTAOBARRA_ID_N & "," & _
                       VENDEDOR_ID_N & "," & _
                       NUMERO_CAIXA_CPU & ",'" & CGCCPF & "'," & USUARIO_ID_N & ",'" & DT_REQ & "','" & _
                       STATUS & "','" & TIPO_REGISTRO & "'," & USUARIO_LIBERA_VENDA & ",'" & _
                       tpMOEDA(Valor_Desconto) & "','" & tpMOEDA(PERC_DESC) & "','" & Replace(NOME_CLIENTE, "'", " ") & "','" & _
                       tpMOEDA(VALOR_RECEBIDO) & "','" & tpMOEDA(VALOR_TOTAL) & "','" & PREFIXO & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spPedido"
End Sub

Public Sub spPedidoItem(Acao_N As Integer, _
                        QTD_PEDIDA As Double, Valor_Item As Double, PERC_DESC As Double, CFOP_ID As String, _
                        STRIBUTARIA As String, VLRBASEICMS As Double, PERCICMS As Double, VLRICMS As Double, _
                        VLRBASEICMSSUBST As Double, PERCICMSSUBST As Double, VLRICMSSUBST As Double, _
                        PERCREDUCAOICMS As Double, PERCIVA As Double, PERC_IPI As Double, VLR_IPI As Double, _
                        Valor_Desconto As Double, STATUS As String, PRECO_CUSTO As Double, TIPO_REG As String, _
                        PESO_ITEM As Double, QTDE_BALANCA As Double, ALTURA As Double, LARGURA As Double, _
                        SEQ_PEDIDO_ID As Long, SEQ_COMANDA_ID As Long, ORIGEM_ITEM_A As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   SQL = "spPedidoItem " & Acao_N & "," & _
                           PEDIDO_ID_N & "," & SEQ_PEDIDO_ID & "," & PRODUTO_ID_N & ",'" & tpMOEDA(QTD_PEDIDA) & "','" & tpMOEDA(Valor_Item) & "','" & tpMOEDA(PERC_DESC) & "','" & _
                           CFOP_ID & "','" & STRIBUTARIA & "','" & tpMOEDA(VLRBASEICMS) & "','" & tpMOEDA(PERCICMS) & "','" & tpMOEDA(VLRICMS) & "','" & _
                           tpMOEDA(VLRBASEICMSSUBST) & "','" & tpMOEDA(PERCICMSSUBST) & "','" & tpMOEDA(VLRICMSSUBST) & "','" & tpMOEDA(PERCREDUCAOICMS) & "','" & tpMOEDA(PERCIVA) & "','" & _
                           tpMOEDA(PERC_IPI) & "','" & tpMOEDA(VLR_IPI) & "','" & tpMOEDA(Valor_Desconto) & "','" & STATUS & "','" & tpMOEDA(PRECO_CUSTO) & "','" & TIPO_REG & "','" & _
                           tpMOEDA(PESO_ITEM) & "','" & tpMOEDA(QTDE_BALANCA) & "'," & ATENDENTE_ID_N & ",'" & tpMOEDA(ALTURA) & "','" & tpMOEDA(LARGURA) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

'MsgBox " ORIGEM_ITEM_A = " & ORIGEM_ITEM_A

 '  If PEDIDO_ID_N > 0 And CARTAOBARRA_ID_N > 0 Then
      'If SEQ_COMANDA_ID <= 0 Then _
         SEQ_COMANDA_ID = 0 & MAX_ID("SEQ_COMANDA_ID", "PEDIDOCOMANDA", "CARTAOBARRA_ID", Str(CARTAOBARRA_ID_N), "", "")

  '    SEQ_COMANDA_ID = 0 & MAX_ID("SEQ_COMANDA_ID", "PEDIDOCOMANDA", "CARTAOBARRA_ID", Str(CARTAOBARRA_ID_N), "", "")

    '  spPedidoComanda ACAO_N, SEQ_COMANDA_ID, SEQ_PEDIDO_ID, ORIGEM_ITEM_A
   'End If

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spPedidoItem"
End Sub


Public Function CONSULTA_EQP_VEICULO() As String
'On Error GoTo ERRO_TRATA

   SQL3 = ""
   CONSULTA_EQP_VEICULO = ""
   If INDR_OS_VEICULO = True Then
      frmOSVeiculoConsulta.Show 1
      Else: frmOSEqpConsulta.Show 1
   End If
   If Trim(SQL3) <> "" Then _
      CONSULTA_EQP_VEICULO = "" & SQL3

   SQL3 = ""

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "CONSULTA_EQP_VEICULO"
End Function

Public Sub spOSTERMO(Acao_N As Integer, OSTERMO_ID_N As Long, OS_ID_N As Long, OSTERMOOBS_A As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   OSTERMOOBS_A = Replace(OSTERMOOBS_A, "'", " ")
   OSTERMOOBS_A = Replace(OSTERMOOBS_A, "'", "")
   OSTERMOOBS_A = Replace(OSTERMOOBS_A, ".", "")
   OSTERMOOBS_A = Replace(OSTERMOOBS_A, "-", "")
   OSTERMOOBS_A = Replace(OSTERMOOBS_A, "/", "")

   SQL = "spOSTERMO " & Acao_N & "," & OSTERMO_ID_N & "," & OS_ID_N & ",'" & Trim(OSTERMOOBS_A) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spOSTERMO"
End Sub

Public Sub spOSOBS(Acao_N As Integer, OSOBS_ID_N As Long, OS_ID_N As Long, OSOBS_A As String)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   OSOBS_A = Replace(OSOBS_A, "'", " ")
   OSOBS_A = Replace(OSOBS_A, "'", "")
   OSOBS_A = Replace(OSOBS_A, ".", "")
   OSOBS_A = Replace(OSOBS_A, "-", "")
   OSOBS_A = Replace(OSOBS_A, "/", "")

   SQL = "spOSOBS " & Acao_N & "," & OSOBS_ID_N & "," & OS_ID_N & ",'" & Trim(OSOBS_A) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spOSOBS"
End Sub

Public Sub spNFITEM(Acao_N As Integer, _
                    NF_ID As Long, SEQ_ID As Long, PRODUTO_ID As Long, VALOR As Double, DESCONTO As Double, _
                    QTDE As Double, CFOP_ID As String, STRIBUTARIA As String, VLRBASEICMS As Double, PERCICMS As Double, _
                    VLRICMS As Double, VLRBASEICMSSUBST As Double, PERCICMSSUBST As Double, VLRICMSSUBST As Double, _
                    PERCREDUCAOICMS As Double, PERCIVA As Double, PERC_IPI As Double)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   SQL = "spNFITEM " & Acao_N & "," & _
                       NF_ID & "," & SEQ_ID & "," & PRODUTO_ID & ",'" & Trim(tpMOEDA(VALOR)) & "','" & Trim(tpMOEDA(DESCONTO)) & "','" & _
                       Trim(tpMOEDA(QTDE)) & "','" & CFOP_ID & "','" & STRIBUTARIA & "','" & Trim(tpMOEDA(VLRBASEICMS)) & "','" & Trim(tpMOEDA(PERCICMS)) & "','" & _
                       Trim(tpMOEDA(VLRICMS)) & "','" & Trim(tpMOEDA(VLRBASEICMSSUBST)) & "','" & Trim(tpMOEDA(PERCICMSSUBST)) & "','" & Trim(tpMOEDA(VLRICMSSUBST)) & "','" & Trim(tpMOEDA(PERCREDUCAOICMS)) & "','" & _
                       Trim(tpMOEDA(PERCIVA)) & "','" & Trim(tpMOEDA(PERC_IPI)) & "'"

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spNFITEM"
End Sub

Public Sub PREPARA_TRIBUTACAO_PRODUTO(CNPJ_CPF_A As String, VLR_UNIT_a As String, QTDE_PEDIDO_a As String)
'On Error GoTo ERRO_TRATA

'UF DO CLIENTE
'IVA
'MVA (Margem de Valor Agregado)
'PEGAR ALIQUOTA PARA O ESTADO TAL

   Dim VLR_UNIT_N    As Double
   Dim QTDE_PEDIDO_N As Double

   VLR_UNIT_N = 0 & Trim(VLR_UNIT_a)
   QTDE_PEDIDO_N = 0 & Trim(QTDE_PEDIDO_a)

   If VLR_UNIT_N <= 0 Then
      MsgBox "Valor não informado."
      Exit Sub
   End If
   If QTDE_PEDIDO_N <= 0 Then
      MsgBox "Quantidade não informado."
      Exit Sub
   End If
   CNPJ_CPF_A = Trim(CNPJ_CPF_A)
   If Trim(CNPJ_CPF_A) = "" Then
      MsgBox "CNPJ/CPF não informado."
      Exit Sub
   End If

   Dim CFOP_ID_N As Integer

   UF_CLIENTE_A = ""

CLIENTE_ID_N = 0 & TRAZ_ID_TABELA("CLIENTE", "cliente_id", "cgccpf", CNPJ_CPF_A)

   If CLIENTE_ID_N < 0 Then
      MsgBox "Cliente não informado, verifique !!!"
      Exit Sub
   End If

   Dim tabEnd  As New ADODB.Recordset

   If tabEnd.State = 1 Then _
      tabEnd.Close

   SQL = "select CEP.UF from CLIENTE WITH (NOLOCK)"
   SQL = SQL & " INNER JOIN ENDERECO WITH (NOLOCK)"
   SQL = SQL & " ON CLIENTE.PESSOA_ID = ENDERECO.PESSOA_ID "
   SQL = SQL & " INNER JOIN CEP WITH (NOLOCK)"
   SQL = SQL & " ON ENDERECO.CEP_ID = CEP.Cep_ID"

   SQL = SQL & " where CLIENTE.cliente_id = " & CLIENTE_ID_N
   SQL = SQL & " and tipo = 'C'"

   tabEnd.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not tabEnd.EOF Then _
      UF_CLIENTE_A = Trim(tabEnd.Fields("UF").Value)
   If tabEndereco.State = 1 Then _
      tabEndereco.Close

   If Trim(CNPJ_CPF_A) <> "99999999999" And Trim(CNPJ_CPF_A) <> "" Then
      TRATA_PESSOA Trim(CNPJ_CPF_A)
      If Trim(UF_CLIENTE_A) = "" Then
         If INDR_PEDIDO_VENDA = False Then
            MsgBox "Cliente com cadastro incompleto !!! UF_CLIENTE_A = " & UF_CLIENTE_A
            Exit Sub
         End If
      End If
   End If

   If Trim(UF_EMPRESA_A) = "" Then _
      PEGA_DADOS_EMPRESA

   'aqui é ajustado se for consumidor final tem que pegar o mesmo UF de destrino para aliquotas
   If Trim(CNPJ_CPF_A) = "99999999999" Then _
      UF_CLIENTE_A = "" & UF_EMPRESA_A

   If Trim(UF_CLIENTE_A) = "" Then _
      UF_CLIENTE_A = "" & UF_EMPRESA_A

   Dim rstProduto          As New ADODB.Recordset
   Dim ST_PRODUTO_A        As String
   Dim PERCIVA_A           As String
   Dim COMP_TRIBUTARIA_A   As String

   ST_PRODUTO_A = ""

   If rstProduto.State = 1 Then _
      rstProduto.Close

   SQL = "select situacao_tributaria,perciva,comp_tributaria from PRODUTO WITH (NOLOCK)"
   SQL = SQL & " where produto_id = " & PRODUTO_ID_N
   rstProduto.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
   If rstProduto.EOF Then
      If rstProduto.State = 1 Then _
         rstProduto.Close

      MsgBox "Rotina Tributação, produto não cadastrado." & vbCrLf & "Verique"
      Exit Sub
      Else
         ST_PRODUTO_A = "" & Trim(rstProduto.Fields("SITUACAO_TRIBUTARIA").Value)
         PERCIVA_A = "" & rstProduto!PERCIVA
         COMP_TRIBUTARIA_A = "" & rstProduto!COMP_TRIBUTARIA
   End If
   If rstProduto.State = 1 Then _
      rstProduto.Close

   Dim TabTempCFOP                  As New ADODB.Recordset
   Dim VALOR_BASE_ICMS_N            As Double
   Dim PERC_ICMS_N                  As Double
   Dim VALOR_BASE_ICMS_SUBST_N      As Double
   Dim VALOR_ICMS_PRODUTO_SUBST_N   As Double
   Dim VALOR_PERC_ICMS_SUBST_N      As Double
   Dim strCFOP_ITEM                 As String
   Dim PERC_REDUCAO_ICMS_N          As Double
   Dim PERC_IVA_N                   As Double
   Dim VALOR_TOTAL_ITEM_N           As Double
   Dim Aliquota_N                   As Double
   Dim VALOR_ICMS_N                 As Double

   VALOR_BASE_ICMS_N = 0
   VALOR_ICMS_N = 0
   PERC_ICMS_N = 0
   VALOR_BASE_ICMS_SUBST_N = 0
   VALOR_ICMS_PRODUTO_SUBST_N = 0
   VALOR_PERC_ICMS_SUBST_N = 0
   PERC_REDUCAO_ICMS_N = 0
   PERC_IVA_N = 0
   Aliquota_N = 0
   VALOR_TOTAL_ITEM_N = 0

   strCFOP_ITEM = ""
   strCFOP_ITEM = "5102"   'CFOP 5102 - Venda de mercadoria adquirida ou recebida de terceiros
   'strCFOP_ITEM = "5101"   'CFOP 5101 - Venda de produção do estabelecimento

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0

   VALOR_ITEM_N = 0 & VLR_UNIT_N
   QTDE_N = 0 & QTDE_PEDIDO_N
   VALOR_TOTAL_ITEM_N = (QTDE_N * VALOR_ITEM_N)

   ALIQUOTA_ICMS_NORMAL_DENTRO_UF = 0
   ALIQUOTA_ICMS_NORMAL_FORA_UF = 0

'BUSCANDO ALIQUOTAS
   If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
      strCFOP_ITEM = "5102"
      Else: strCFOP_ITEM = "6102"
   End If

   'Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, UF_CLIENTE_A, 0)
   If IsNumeric(strCFOP_ITEM) Then _
      CFOP_ID_N = 0 & strCFOP_ITEM
   Call BUSCA_ALIQUOTA_ICMS(UF_EMPRESA_A, "", CFOP_ID_N)

   '5405  Venda de mercadoria, adquirida ou recebida de terceiros,
   'sujeita ao regime de substituição tributária,
   'na condição de contribuinte-substituído

   'Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros
   'em operação com mercadorias sujeitas ao regime de substituição tributária,
   'na condição de contribuinte substituído.
   'strCFOP_ITEM = "5405"   'não é industria


'28/03/2017 VERIFICAR SE É ASSIM MESMO:
'QUANDO CLIENTE É CONSUMIDOR FINAL NÃO PASSA NO SEFAZ O PRODUTO COMO SUBSTITUIÇÃO TRIBUTÁRIA
'DAI MUDO AQUI MANUALMENTE A ST DO ITEM PARA 00-TRIBUTADO INTEGRALMENTE
'SIMPLES NACIONAL
If CTR_EMPRESA_N = 1 Then _
   If Trim(UCase(CCE_CLIENTE_A)) = "ISENTO" Or Trim(CCE_CLIENTE_A) = "" Then _
      ST_PRODUTO_A = "00"


'TEM QUE VER QUANDO NÃO FOR VENDA

   Select Case ST_PRODUTO_A
      Case "00"   'Tributada integralmente
         'If CTR_EMPRESA_N = 1 Then  'se é optante do simples nacional
            If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
               strCFOP_ITEM = "5102"
               Else: strCFOP_ITEM = "6102"
            End If
         'End If

         If INDR_INDUSTRIA_B = True Then  'se produto o proprio produto pra vender
            'Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros,
            'na condição de contribuinte substituto,
            'em operação com mercadorias sujeitas ao regime de substituição tributária.
            strCFOP_ITEM = "5403"   'CFOP 5403 - Venda de mercadoria adquirida ou recebida de terceiros
                                    'em operação com mercadoria sujeita ao regime de substituição tributária,
                                    'na condição de contribuinte substituto

            If CTR_EMPRESA_N = 1 Then  'se é optante do simples nacional
               If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
                  'INDR_PRODUTO_PRODUCAO_B = vem da tabela de familia informando que o produto é de produção
                  If INDR_PRODUTO_PRODUCAO_B = True Then _
                     strCFOP_ITEM = "5101"   'CFOP 5101 - Venda de produção do estabelecimento
                  Else
                     strCFOP_ITEM = "6102"
                     If INDR_PRODUTO_PRODUCAO_B = True Then _
                        strCFOP_ITEM = "6101"   'CFOP 6101 - Venda de produção do estabelecimento
               End If
            End If
         End If   'If INDR_INDUSTRIA_B = True Then

         'Desconto nao entra no valor do ICMS de acordo com informacoes da CONTABILIDADE
         VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N

         '==================EU
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then       'DENTRO DO ESTADO ICMS NORMAL
            PERC_ICMS_N = ALIQUOTA_ICMS_NORMAL_DENTRO_UF
            Else: PERC_ICMS_N = ALIQUOTA_ICMS_NORMAL_FORA_UF   'FORA DO ESTADO ICMS NORMAL
         End If
         VALOR_ICMS_N = ((VALOR_BASE_ICMS_N * PERC_ICMS_N) / 100)
      Case "10"   'Tributada  e com cobrança do ICMS por substituição tributária
         VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
   
         Aliquota_N = ALIQUOTA_ICMS_NORMAL_DENTRO_UF
         If Trim(UF_CLIENTE_A) <> Trim(UF_EMPRESA_A) Then _
            Aliquota_N = ALIQUOTA_ICMS_NORMAL_FORA_UF
   
         If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then
            'Campo IVA nao existe nao tabela verificar se precisa, Índices de Valor Agregado
            If Not IsNull(PERCIVA_A) Then _
              VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * PERCIVA_A) / 100)  'Valor da Reducao da base
   
            'VALOR_BASE_ICMS_SUBST_N = ((VALOR_BASE_ICMS_N * 1) / 100)  'Valor da Reducao da base
            VALOR_ICMS_PRODUTO_SUBST_N = ((VALOR_BASE_ICMS_SUBST_N * Aliquota_N) / 100)  'é fixo o percentual, procurar saber se tem como parametrizar
            VALOR_PERC_ICMS_SUBST_N = Aliquota_N
         End If
      Case "20"   'Com redução de base de cálculo
         If COMP_TRIBUTARIA_A = 0 Then 'tipos de maquinas, normais, agricolas, industriais
            If CCE_CLIENTE_A <> "" Then    'Tem que ter inscricao estadual
               VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CONTRIB) / 100)
               PERC_REDUCAO_ICMS_N = TP2_DE_CONTRIB
               Else  'Sem inscricao estadual
                  VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NCONTRIB) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_NCONTRIB
            End If
         End If
   
         'Maquinas agricolas
         If COMP_TRIBUTARIA_A = 1 Then
            If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
               If CCE_CLIENTE_A <> "" Then
                  VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CMAQ_IMP) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_CMAQ_IMP
                  Else
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NMAQ_IMP) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_DE_NMAQ_IMP
               End If
               Else 'Fora do Estado
                  If CCE_CLIENTE_A <> "" Then
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_CMAQ_IMP) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_FE_CMAQ_IMP
                     Else
                        VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_NMAQ_IMP) / 100)
                        PERC_REDUCAO_ICMS_N = TP2_FE_NMAQ_IMP
                  End If
            End If
         End If
   
         If COMP_TRIBUTARIA_A = 2 Then 'Maquinas industriais
            If Trim(UF_CLIENTE_A) = Trim(UF_EMPRESA_A) Then 'Dentro do estado
               If CCE_CLIENTE_A <> "" Then
                  VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_CONTRIB) / 100)
                  PERC_REDUCAO_ICMS_N = TP2_DE_CONTRIB
                  Else
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_DE_NCONTRIB) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_DE_NCONTRIB
               End If
               Else 'Fora do Estado
                  If CCE_CLIENTE_A <> "" Then
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_CAP_INDU) / 100)
                     PERC_REDUCAO_ICMS_N = TP2_FE_CAP_INDU
                     Else
                        VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TP2_FE_NAP_INDU) / 100)
                        PERC_REDUCAO_ICMS_N = TP2_FE_NAP_INDU
                  End If
            End If
         End If
      Case "30"   'Isenta ou não tributada e com cobrança do ICMS por substituição tributária
         VALOR_BASE_ICMS_N = 0
         VALOR_ICMS_N = 0
         PERC_ICMS_N = 0
   
         If UCase(UF_CLIENTE_A) <> UCase(UF_EMPRESA_A) Then
             '//Desconto nao entra no valor de ICMS de Acordo com as
             '//Informacoes Contabeis
             '//move (ITENS.TOTAL_ITEM - ITENS.VLR_DESC_RATEIO)  ;
             '//                                     To   ITENS.VLR_BASE_ICMS
             VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
             '??? nao grava o percentual do aliquota?
         End If
      Case "40"   '40 Isenta
         VALOR_BASE_ICMS_N = 0
         VALOR_ICMS_N = 0
         PERC_ICMS_N = 0
      Case "41"   'Não tributada
         VALOR_BASE_ICMS_N = 0
         VALOR_ICMS_N = 0
         PERC_ICMS_N = 0
      Case "50"   'Suspensão
      Case "51"   'Diferimento
      Case "60"   'ICMS cobrado anteriormente por substituição tributária
         VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
         If UCase(UF_CLIENTE_A) = UCase(UF_EMPRESA_A) Then
            If TIPO_CLIENTE_N = 2 Then 'Atacado
               '//Dentro do Estado e Cliente Contribuinte ele e Isento
               '/Emanoel Informacoes Contabilidade dia 30/05/2006
               VALOR_BASE_ICMS_N = 0
               VALOR_ICMS_N = 0
               PERC_ICMS_N = 0
            End If
            'Só é tratado o tipo de cliente 2, atacado, e os outros tipos de clientes (varejo),
            'nao precisa tratar?
            Else 'Fora do estado
               If TIPO_CLIENTE_N = 2 Then 'Atacado
                  VALOR_BASE_ICMS_N = VALOR_TOTAL_ITEM_N
                  'nao grava o percentual? porque?
               End If
         End If

         'DENTRO DO ESTADO
         If UCase(UF_CLIENTE_A) = UCase(UF_EMPRESA_A) Then
            'If Trim(ST_PRODUTO_A) = 60 Then
               'CFOP 5102 - Venda de mercadoria adquirida ou recebida de terceiros
               'CFOP 5405 - Venda de mercadoria adquirida/recebida de terceiros em operação _
                            com mercadoria sujeita ao regime de substituição tributária, na condição de _
                            contrib substituído
       
      'portanto o que vai diferenciar se será um codigo ou outro será a mercadoria em
      'si...se ela é substituiçao tributaria ou nao...se for varias mercadorias vc tem que
      'verificar uma por uma pra saber.
      
               strCFOP_ITEM = "5405"
               'Else: strCFOP_ITEM = CFOP_SAIDA_DENTRO_UF_N                     'cfop de venda dentro do estado
            'End If   'If Trim(ST_PRODUTO_A) = 60 Then
         End If
      
         'FORA DO ESTADO
         If UCase(UF_CLIENTE_A) <> UCase(UF_EMPRESA_A) Then
            'If Trim(ST_PRODUTO_A) = 60 Then
               strCFOP_ITEM = "6403"  'Fixo por enquanto
               '6403 Venda de mercadoria adquirida ou recebida de terceiros em operação _
                     com mercadoria sujeita ao regime de substituição tributária, _
                     na condição de contribuinte substituto _
                     Classificam-se neste código as vendas de mercadorias adquiridas ou recebidas de terceiros, _
                     na condição de contribuinte substituto, em operação com mercadorias sujeitas _
                     ao regime de substituição tributária.
      
               strCFOP_ITEM = "6404"
               '6404 Venda de mercadoria sujeita ao regime de substituição tributária, _
                     cujo imposto já tenha sido retido anteriormente _
                     Classificam-se neste código as vendas de mercadorias sujeitas ao regime de substituição tributária, _
                     na condição de substituto tributário, exclusivamente nas hipóteses em que o _
                     imposto já tenha sido retido anteriormente
      
            '   Else: strCFOP_ITEM = CFOP_SAIDA_FORA_UF_N                  'cfop de venda fora do estado do estado
            'End If
      
            SQL = "select * from CFOP WITH (NOLOCK)"
            SQL = SQL & " Where CFOP_ID = '" & Trim(strCFOP_ITEM) & "'"
            TabTempCFOP.Open SQL, CONECTA_RETAGUARDA, , , adCmdText
            If TabTempCFOP.EOF Then
               If TabTempCFOP.State = 1 Then _
                  TabTempCFOP.Close
      
               MsgBox "O sistema não localizou o CFOP de numero=" & strCFOP_ITEM & vbCrLf & "Não é possivel continuar a processar"
               'fazer procedimento de reverter ou entao, deixar a pessoa processar novamente. Verificar o melhor
               Exit Sub
               Else
                  If Trim(Len(CNPJCPF_A)) > 11 Then ' Se for pessoa juridica
                     VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TabTempCFOP!ALIQUOTA_ICMS_DENTRO) / 100) 'CFOP.P_ICMS_VND_F_UF - verificar se existe
                     PERC_ICMS_N = TabTempCFOP!ALIQUOTA_ICMS_DENTRO ' CFOP.P_ICMS_VND_F_UF'duas aliquotas para  o mesmo cfop
                     Else ' Pessoa fisica
                        VALOR_ICMS_N = ((VALOR_TOTAL_ITEM_N * TabTempCFOP!ALIQUOTA_ICMS_FORA) / 100)
                        PERC_ICMS_N = TabTempCFOP!ALIQUOTA_ICMS_FORA
                  End If
            End If
            If TabTempCFOP.State = 1 Then _
               TabTempCFOP.Close
         End If
      Case "70"   'Com redução de base de cálculo e cobrança de ICMS por substituição tributária
      Case "90"   'Outras
   End Select

   If VALOR_BASE_ICMS_N = 0 Then
      PERC_ICMS_N = 0
      VALOR_ICMS_N = 0
   End If

'ATUALIZAR PEDIDO
   SQL = "UPDATE PEDIDOITEM SET "

   SQL = SQL & " VlrBaseIcms = " & tpMOEDA(VALOR_BASE_ICMS_N)
   SQL = SQL & ", PERCICMS = " & tpMOEDA(PERC_ICMS_N)
   SQL = SQL & ", VlrIcms = " & tpMOEDA(VALOR_ICMS_N)

   SQL = SQL & ", VLRBASEICMSSUBST = " & tpMOEDA(VALOR_BASE_ICMS_SUBST_N)
   SQL = SQL & ", PERCICMSSUBST = " & tpMOEDA(VALOR_PERC_ICMS_SUBST_N)
   SQL = SQL & ", VLRICMSSUBST = " & tpMOEDA(VALOR_ICMS_PRODUTO_SUBST_N)

   SQL = SQL & ", cfop_id = '" & Trim(strCFOP_ITEM) & "'"
   SQL = SQL & ", STRIBUTARIA = '" & ST_PRODUTO_A & "'"

   SQL = SQL & " Where pedido_id = " & PEDIDO_ID_N
   SQL = SQL & " and produto_id = " & PRODUTO_ID_N

   CONECTA_RETAGUARDA.Execute SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGERAL", "PREPARA_TRIBUTACAO_PRODUTO"
End Sub

Public Sub spPEDIDOFATURA(Acao_N As Integer, PEDIDOFATURA_ID As Long, PEDIDO_ID As Long, TABELAPRECO_ID As Integer, FORMAPAGTO_ID As Integer, TIPOVENDA_ID As Long)
'On Error GoTo ERRO_TRATA

   If Acao_N <= 0 Then
      MsgBox "Informar Acao_N de instrução para procedimento."
      Exit Sub
   End If

   SQL = "spPEDIDOFATURA " & Acao_N & "," & PEDIDOFATURA_ID & "," & PEDIDO_ID & "," & TABELAPRECO_ID & "," & FORMAPAGTO_ID & "," & TIPOVENDA_ID

   CONECTA_RETAGUARDA.Execute "EXEC " & SQL

Exit Sub
ERRO_TRATA:
   TRATA_ERROS Err.Description, "mdlGeral", "spPEDIDOFATURA"
End Sub
