Attribute VB_Name = "mdlNFce"
Public Sub TESTA_FLEX_DOC()
'
'   Exemplo para obter vers�o da DLL em uso
'

'
' instancia classe
'
Dim objCTeUtil As CTe_Util.Util
MsgBox Application.Path
Set objCTeUtil = New CTe_Util.Util
 
'
' obtem vers�o
'

MsgBox "A vers�o da DLL �: " + objCTeUtil.Versao, vbInformation, "Resultado"
'
' libera classe
'
Set objCTeUtil = Nothing
End Sub
