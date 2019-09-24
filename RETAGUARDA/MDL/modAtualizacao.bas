Attribute VB_Name = "modAtualizacao"
Option Explicit
   Dim strSQL  As String
   Dim Rs      As New ADODB.Recordset
   Dim Rs2     As New ADODB.Recordset

Public Sub VerificaCampo(strTabela As String, strColuna As String, strTipo As String, Optional tamanho As Long, Optional ValorPadrao As String, Optional strDescricao As String)
'On Error GoTo ERRO_TRATA

   Dim INTRETORNO As Long
   Dim sSQL       As String

    Rs.Open "select ColunaCerta.length as tamanho, systypes.name as tipo from sysobjects TabelaCerta INNER JOIN syscolumns ColunaCerta ON TabelaCerta.id = ColunaCerta.id INNER JOIN systypes ON ColunaCerta.xtype = systypes.xtype WHERE (TabelaCerta.xtype IN ('U', 'V')) AND (TabelaCerta.name = '" & strTabela & "') AND (ColunaCerta.name = '" & strColuna & "')", CONECTA_RETAGUARDA, , , adCmdText
    If Not Rs.EOF Then
        
        If Rs!tamanho <> tamanho And (Trim(UCase(strTipo)) = "VARCHAR" Or (Trim(UCase(strTipo)) = "CHAR")) Then
            sSQL = "ALTER TABLE " & strTabela & " ALTER COLUMN " & strColuna & " " & strTipo & "(" & tamanho & ")"
            CONECTA_RETAGUARDA.Execute sSQL
        ElseIf UCase(Trim(strTipo)) <> UCase(Trim(Rs!TIPO)) Then
            CONECTA_RETAGUARDA.Execute "ALTER TABLE " & strTabela & " ALTER COLUMN " & strColuna & " " & strTipo
        End If
        If Trim(ValorPadrao) <> "" Then
            sSQL = "ALTER TABLE " & strTabela & " DROP CONSTRAINT DF_" & strTabela & "_" & strColuna
            CONECTA_RETAGUARDA.Execute sSQL
            sSQL = "ALTER TABLE " & strTabela & " ADD CONSTRAINT DF_" & strTabela & "_" & strColuna & " DEFAULT " & ValorPadrao & " FOR " & strColuna
            CONECTA_RETAGUARDA.Execute sSQL
        End If
        If Trim(strDescricao) <> "" Then
            CONECTA_RETAGUARDA.Execute "DECLARE @v sql_variant SET @v = N'" & strDescricao & "' EXECUTE sp_addextendedproperty N'MS_Description', @v, N'user', N'dbo', N'table', N'" & strTabela & "', N'column', N'" & strColuna & "'"
            CONECTA_RETAGUARDA.Execute "DECLARE @v sql_variant SET @v = N'" & strDescricao & "' EXECUTE sp_updateextendedproperty N'MS_Description', @v, N'user', N'dbo', N'table', N'" & strTabela & "', N'column', N'" & strColuna & "'"
        End If
    Else
        'cria campo
        If (Trim(UCase(strTipo)) = "VARCHAR") Or (Trim(UCase(strTipo)) = "CHAR") Then
            sSQL = "ALTER TABLE " & strTabela & " ADD " & strColuna & " " & strTipo & "(" & tamanho & ")"
            CONECTA_RETAGUARDA.Execute sSQL
        Else
            CONECTA_RETAGUARDA.Execute "ALTER TABLE " & strTabela & " ADD " & strColuna & " " & strTipo
        End If
        
        If Trim(strDescricao) <> "" Then
            CONECTA_RETAGUARDA.Execute "DECLARE @v sql_variant SET @v = N'" & strDescricao & "' EXECUTE sp_addextendedproperty N'MS_Description', @v, N'user', N'dbo', N'table', N'" & strTabela & "', N'column', N'" & strColuna & "'"
        End If
        
        If Trim(ValorPadrao) <> "" Then
            CONECTA_RETAGUARDA.Execute "ALTER TABLE " & strTabela & " ADD CONSTRAINT DF_" & strTabela & "_" & strColuna & " DEFAULT " & ValorPadrao & " FOR " & strColuna
        End If
    End If
    Rs.Close
    
    Exit Sub
ERRO_TRATA:
    If Err.Number = "-2147217900" Then 'se não existir a chave
        Resume Next
    End If
    INTRETORNO = MsgBox("VerificaEXISTE_CAMPO_TABELA: Erro ao atualizar banco de dados" & Err.Number & " - " & Err.Description & "  Origem: " & Err.Source, 50, "Modulo de Atualização")
    If INTRETORNO = 3 Then
        Rs.Close
        Exit Sub 'Sai da rotina
    ElseIf INTRETORNO = 4 Then
        Resume  'Volta a Linha Atual
    ElseIf INTRETORNO = 5 Then
        Resume Next 'Vai p/ próxima linha
    End If
End Sub

Public Function EXISTE_CAMPO_TABELA(strBanco As String, strCampo As String, strTabela As String) As Boolean
'On Error GoTo ERRO_TRATA

   Dim rsCampo As New ADODB.Recordset
   EXISTE_CAMPO_TABELA = True

   If rsCampo.State = 1 Then _
      rsCampo.Close

   SQL = "select isnull(count(syscolumns.name),0) as qtde from sysobjects "
   SQL = SQL & " INNER JOIN syscolumns "
   SQL = SQL & " ON sysobjects.id = syscolumns.id "
   SQL = SQL & " WHERE (sysobjects.xtype = 'U') "
   SQL = SQL & " and sysobjects.name = '" & strTabela & "' "
   SQL = SQL & " and syscolumns.name = '" & strCampo & "'"

   If Trim(strBanco) = "RETAGUARDA" Then _
      rsCampo.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Trim(strBanco) = "SHFINFO" Then _
      rsCampo.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Trim(strBanco) = "GLOBAL" Then _
      rsCampo.Open SQL, CONECTA_GLOBAL, , , adCmdText

   If Not rsCampo.EOF Then
      If rsCampo!QTDE = 0 Then
         EXISTE_CAMPO_TABELA = False
         Else: EXISTE_CAMPO_TABELA = True
      End If
      Else: EXISTE_CAMPO_TABELA = False
   End If
   If rsCampo.State = 1 Then _
      rsCampo.Close

Exit Function
ERRO_TRATA:
    EXISTE_CAMPO_TABELA = False
   If rsCampo.State = 1 Then _
      rsCampo.Close
End Function

Public Function EXISTE_OBJ_BANCO(strBanco As String, strTabela As String, TIPO_OBJ As String) As Boolean
'On Error GoTo ERRO_TRATA

   Dim rsTabela As New ADODB.Recordset
   EXISTE_OBJ_BANCO = True

   If rsTabela.State = 1 Then _
      rsTabela.Close

   If Trim(strBanco) = "" Then _
      Exit Function

   SQL = "select isnull(count(sysobjects.name),0) AS qtde from sysobjects "
   SQL = SQL & " WHERE sysobjects.name = '" & strTabela & "'"

   If Trim(TIPO_OBJ) <> "" Then _
      SQL = SQL & " and sysobjects.type = '" & Trim(TIPO_OBJ) & "'"

   'rsTabela.Open "select isnull(count(sysobjects.name),0) AS qtde from sysobjects WHERE (sysobjects.xtype = 'U') and sysobjects.name = '" & strTabela & "'", CONECTA_RETAGUARDA, , , adCmdText

   If Trim(strBanco) = "RETAGUARDA" Then _
      rsTabela.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Trim(strBanco) = "SHFINFO" Then _
      rsTabela.Open SQL, CONECTA_RETAGUARDA, , , adCmdText

   If Trim(strBanco) = "GLOBAL" Then _
      rsTabela.Open SQL, CONECTA_GLOBAL, , , adCmdText

   If Not rsTabela.EOF Then
      If rsTabela!QTDE = 0 Then
         EXISTE_OBJ_BANCO = False
         Else: EXISTE_OBJ_BANCO = True
      End If
      Else: EXISTE_OBJ_BANCO = False
   End If
   If rsTabela.State = 1 Then _
      rsTabela.Close

Exit Function
ERRO_TRATA:
   EXISTE_OBJ_BANCO = False
   rsTabela.Close
End Function

Public Function ExisteIndice(strIndice As String, strTabela As String) As Boolean
'horacio 26/03/2009 11:33
'On Error GoTo ERRO_TRATA

   Dim rsIndice   As New ADODB.Recordset
   Dim sSQL       As String

   ExisteIndice = True

   sSQL = "select * from sys.indexes "
   sSQL = sSQL & " WHERE name = '" & strIndice & "'"
   rsIndice.Open sSQL, CONECTA_RETAGUARDA, , , adCmdText
   If Not rsIndice.EOF Then
      If rsIndice!QTDE = 0 Then
         ExisteIndice = False
         Else: ExisteIndice = True
      End If
      Else: ExisteIndice = False
   End If
   rsIndice.Close

Exit Function
ERRO_TRATA:
   ExisteIndice = False
   rsIndice.Close
End Function

Public Function ExisteProcedure(strProcedure As String) As Boolean
    'On Error GoTo ERRO_TRATA
    Dim rsCampo As New ADODB.Recordset
    ExisteProcedure = True
    
    rsCampo.Open "select COUNT(*) AS qtde from sysobjects WHERE id = object_id(N'[dbo].[" & strProcedure & "]') AND OBJECTPROPERTY(id, N'IsProcedure') = 1", CONECTA_RETAGUARDA, , , adCmdText
    If Not rsCampo.EOF Then
        If rsCampo!QTDE = 0 Then
            ExisteProcedure = False
        Else
            ExisteProcedure = True
        End If
    Else
        ExisteProcedure = False
    End If
    rsCampo.Close
    
    Exit Function
    
ERRO_TRATA:
    ExisteProcedure = False
    rsCampo.Close
End Function

