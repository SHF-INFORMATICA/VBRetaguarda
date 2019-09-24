VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function PROCESSA_TAB_PESSOA(TIPO_CHAMADA_A As String, _
                                    PESSOA_ID_N As Long, _
                                    CNPJCPF_A As String, _
                                    DESCRICAO_A As String, _
                                    RAZAO_A As String, _
                                    DATA_CAD_D As Date, _
                                    SITUACAO_A As String) _
                                    As Boolean
'On Error GoTo ERRO_TRATA

   GRAVA_TAB_PESSOA = False

   If Trim(TIPO_CHAMADA_A) = "" Then
      MsgBox "Parâmetro tipo_chamada_a não informado."
      Exit Function
   End If
   If Trim(CNPJCPF_A) = "" Then
      MsgBox "Parâmetro DESCRICAO_A não informado."
      Exit Function
   End If
   If Trim(DESCRICAO_A) = "" Then
      MsgBox "Parâmetro DESCRICAO_A não informado."
      Exit Function
   End If
   If Trim(RAZAO_A) = "" Then
      MsgBox "Parâmetro RAZAO_A não informado."
      Exit Function
   End If
   If Trim(SITUACAO_A) = "" Then
      MsgBox "Parâmetro SITUACAO_A não informado."
      Exit Function
   End If

   If Trim(UCase(TIPO_CHAMADA_A)) = UCase("inclusao") Then
      If Not IsDate(DATA_CAD_D) Then
         MsgBox "Parâmetro DATA_CAD_D não informado."
         Exit Function
      End If

      CONECTA_RETAGUARDA.Execute "EXEC sp_INSERT_PESSOA " & "'dbo.NOTAENTRADA.CODG_FORNEC'" & "," & "'FORNECEDOR_ID'" & "," & "'COLUMN'"
   End If
   If Trim(UCase(TIPO_CHAMADA_A)) = UCase("alteracao") Then
      If PESSOA_ID_N < 0 Then
         MsgBox "Parâmetro PESSOA_ID_N não informado."
         Exit Function
      End If

      'CONECTA_RETAGUARDA.Execute "EXEC sp_rename " & "'dbo.NOTAENTRADA.CODG_FORNEC'" & "," & "'FORNECEDOR_ID'" & "," & "'COLUMN'"
   End If

Exit Function
ERRO_TRATA:
   TRATA_ERROS Err.Description, "cCadastro", "PROCESSA_TAB_PESSOA"
End Function
