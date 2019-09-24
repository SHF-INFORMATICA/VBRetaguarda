Attribute VB_Name = "Module1"
DefInt A-Z
Option Explicit
'//Define os parametros para as ações em EncryptString
Public Const ENCRYPT = 1, DECRYPT = 2
'---------------------------------------------------------------------
' EncryptString
'---------------------------------------------------------------------
Public Function EncryptString(UserKey As String, Text As String, Action As Single) As String
    
    'define as variaveis usadas
    Dim UserKeyX As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String
    
    '//Obtem os caracteres da chave do usuário
    'define o comprimento da chave do usuario usada na criptografia
    n = Len(UserKey)
    
    'redimensiona o array para o tamanho definido
    ReDim userKeyASCIIS(1 To n)
    
    'preenche o array com caracteres asc
    'Debug.Print UserKey; "=> ";
    For i = 1 To n
        userKeyASCIIS(i) = Asc(Mid$(UserKey, i, 1))
        'Debug.Print userKeyASCIIS(i); " ";
    Next
        
    '//redimensiona o array com o tamanho do texto
    'obtem o caractere de texto
    ReDim TEXTAsciis(Len(Text)) As Integer
    
    'preenche o array com caracteres asc
    'Debug.Print
    'Debug.Print Text; " => ";
    For i = 1 To Len(Text)
        TEXTAsciis(i) = Asc(Mid$(Text, i, 1))
        'Debug.Print TEXTAsciis(i); " ";
    Next
    
    '//cifra/decifra
    If Action = ENCRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           Temp = TEXTAsciis(i) + userKeyASCIIS(j)
           If Temp > 255 Then
              Temp = Temp - 255
           End If
           'Debug.Print Temp; " ";
           rtn = rtn + Chr$(Temp)
           'Debug.Print rtn
       Next
    ElseIf Action = DECRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           Temp = TEXTAsciis(i) - userKeyASCIIS(j)
           If Temp < 0 Then
              Temp = Temp + 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    End If
    
    '//Retorna o texto
    EncryptString = rtn
End Function
