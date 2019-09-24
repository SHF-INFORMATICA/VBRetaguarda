VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNotaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function RetornaTagXML(strCaminhoXML As String, Tag As String) As String
On Error Resume Next
Dim objNodeList, objNode As Object
Dim XML As DOMDocument
Set XML = New DOMDocument

    RetornaTagXML = ""
    XML.async = False
    If XML.Load(strCaminhoXML) Then
        ' *** Tentar pegar o strCampoXML
        Set objNodeList = XML.getElementsByTagName(Tag)
        Set objNode = objNodeList.nextNode
        Dim sLeitura As String
        sLeitura = objNode.Text
        If Len(Trim(sLeitura)) > 0 Then 'Conseguiu ler o XML
            RetornaTagXML = sLeitura
        End If
        Else
        MsgBox "Não foi possível abrir o arquivo XML da NFe especificada para Leitura.", vbCritical, "Erro."
    End If
End Function
