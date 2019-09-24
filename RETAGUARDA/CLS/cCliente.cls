VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************
'Criação: Yuri Grandinetti Lemes
'Propósito: Classe de controle referente a cliente
'*****************************************************************************
Option Explicit

'*****************************************************************************
'Propósito: Verifica se o cliente é optante do Super Simples função so para
'           empresas que trabalha no regime do TARE
'*****************************************************************************
Public Function BuscaDadosdoClienteSuperSimples(lngCodigoCliente As Long, ByRef curPorcIcmsSubst As Currency) As Boolean
'On Error GoTo Err_ClienteSuperSimples
Dim gb_RecordsetII As ADODB.Recordset

    BuscaDadosdoClienteSuperSimples = False
    If g_trabalhacomtare_empresa = 1 Then  ' SÓ VAI ENTRAR AQUI SE A EMPRESA TIPO A VI OU ENESSE FOR OPTANTE DO TARE
       gb_RecordsetII.Source = "SELECT Empresa_Optante_Simples, aliquota_subst_super_simples from cliente WHERE codigo = '" & lngCodigoCliente & "'"
       gb_RecordsetII.Open
    
       ' Set gb_RecordsetII = Conexao.GeraRecordset("SELECT super_simples, aliquota_subst_super_simples " & _
                                                   "from cliente " & _
                                                   "WHERE codigo = '" & lngCodigoCliente & "'", 0)
        If gb_RecordsetII.RecordCount > 0 Then
            If IsNumeric(gb_RecordsetII!aliquota_subst_super_simples) Then
                curPorcIcmsSubst = gb_RecordsetII!aliquota_subst_super_simples
            Else
                curPorcIcmsSubst = 0
            End If
            If gb_RecordsetII!Empresa_Optante_Simples = "1" Then BuscaDadosdoClienteSuperSimples = True
        End If
        gb_RecordsetII.Close
    End If
    Set gb_RecordsetII = Nothing
    
Exit Function
'Err_ClienteSuperSimples: ValidaErros Err, "cCliente - ClienteSuperSimples"
End Function

