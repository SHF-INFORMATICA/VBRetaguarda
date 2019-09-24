Attribute VB_Name = "Module2"
   Public intCodigoContaCorrente As Long
   
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
    rsNossoNumero.Open "select CodigoCarteira, isnull(CodigoBanco,0) as banco, isnull(Agencia,'') as Agencia, isnull(Conta,'') as Conta, isnull(COMRegistro,1) as COMRegistro, isnull(SequenciaBoleta, 0) as SequenciaBoleta, Convenio, ISNULL(SequenciaBoletaFinal,0) AS SequenciaBoletaFinal from ContaCorrente WHERE Empresa_id = " & EMPRESA_ID_N & " AND CodigoContaCorrente = " & intCodigoContaCorrente, CONECTA_RETAGUARDA, , , adCmdText
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
            MsgBox "Atenção. Sequencia do boleto final foi superado, favor entrar em contato com o FINANCEIRO parar solicitar junto ao banco esta nova faixa numérica.", vbCritical, "Calcula Nosso Número"
            CalculaNossonumero = ""
            Exit Function
        End If
        If lngSequenciaBoletoContaCorrente + 500 >= lngSequenciaBoletoFinal Then
            MsgBox "Atenção. Sequencia do boleto final está esgotando, favor entrar em contato com o FINANCEIRO parar solicitar junto ao banco esta nova faixa numérica. Faltam apenas " & lngSequenciaBoletoFinal - lngSequenciaBoletoContaCorrente & " boletos para finalizar e ser bloqueado.", vbCritical, "Calcula Nosso Número"
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
      
    ElseIf intBanco = 104 Then 'CAIXA ECONÔMICA FEDERAL CEF
    
    '            8.2.1 Para a carteira 11 - Cobrança Simples (Vide Nota 3): Número gerado e atribuído pelo sistema de cobrança da
    '            CAIXA para controle interno, e será composto da seguinte forma:
    '            NNNNNNNNNND , onde
    '            NNNNNNNNNN = Número Sequencial
    '            D = Dígito Verificador (calculado pelo Mod. 11)
    '            Obs: para clientes que possuem sistema próprio, preencher o campo com zeros.
    '            8.2.2 Para a carteira 12 - Cobrança Rápida: Número informado pelo cliente, composto da seguinte forma:
    '            9NNNNNNNNND, onde 9 = Fixo
    '            NNNNNNNNN = Número Sequencial
    '            D = Dígito Verificador (calculado pelo Mod. 11)
    
        If intCodigoCarteira = 11 Then 'COBRANÇA SIMPLES
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
            
        ElseIf intCodigoCarteira = 12 Then 'COBRANÇA RÁPIDA
            CalculaNossonumero = "9" & Format(lngSequenciaBoletoContaCorrente, "000000000")
            
        ElseIf intCodigoCarteira = 14 Then 'COBRANÇA REGISTRADA COM CEDENTE
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
        ElseIf intCodigoCarteira = 24 Then 'COBRANÇA SEM REGISTRO COM CEDENTE
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
On Error GoTo ERRO_TRATA

    Dim vInicio As Variant
    
    vInicio = Time
    While DateDiff("s", vInicio, Time) < iSegundos
    Wend
    
    Exit Sub
ERRO_TRATA:
    If Err.number <> 0 Then
        MsgBox "Erro nr: " & Err.number & " Descrição: " & Err.description & " Origem:" & Err.Source, 48, "Tempo"
    End If
End Sub


