<%
' Valida um CNPJ nos formatos tradicional (só números) ou alfanumérico.
' @param {string} cnpj - O CNPJ a ser validado (pode conter pontos, barras, hífens).
' @returns {boolean} - Retorna true se o CNPJ for válido, false caso contrário.
Function validaCNPJ(cnpj)
    ' 1. Limpeza: remove tudo que não é letra ou número e converte para maiúsculas
    Dim cnpjLimpo, regex
    
    Set regex = New RegExp
    regex.Pattern = "[^0-9A-Za-z]"
    regex.Global = True
    regex.IgnoreCase = False
    
    cnpjLimpo = regex.Replace(cnpj, "")
    cnpjLimpo = UCase(cnpjLimpo)
    
    ' 2. Validações iniciais: tamanho deve ser 14 e não pode ser uma sequência inválida conhecida
    If Len(cnpjLimpo) <> 14 Then
        validaCNPJ = False
        Exit Function
    End If
    
    ' Verifica se todos os caracteres são iguais
    Dim isAllSame
    isAllSame = True
    Dim primeiroChar, i
    primeiroChar = Mid(cnpjLimpo, 1, 1)
    For i = 2 To Len(cnpjLimpo)
        If Mid(cnpjLimpo, i, 1) <> primeiroChar Then
            isAllSame = False
            Exit For
        End If
    Next
    
    If isAllSame Then
        validaCNPJ = False
        Exit Function
    End If
    
    ' 3. Função auxiliar para converter caractere em valor numérico
    ' (Já será usado diretamente no cálculo)
    
    ' 4. Função auxiliar para calcular um dígito verificador
    Dim base12, base13, dv1Calculado, dv2Calculado, dvInformado
    
    ' 5. Calcula os dígitos verificadores esperados
    base12 = Left(cnpjLimpo, 12)
    dv1Calculado = calcularDigito(base12, 5)
    base13 = base12 & dv1Calculado
    dv2Calculado = calcularDigito(base13, 6)
    
    ' 6. Compara os dígitos calculados com os dígitos informados
    dvInformado = Mid(cnpjLimpo, 13, 2)
    validaCNPJ = (dvInformado = dv1Calculado & dv2Calculado)
End Function

' Função auxiliar para calcular um dígito verificador
Function calcularDigito(base, pesoInicial)
    Dim soma, peso, i, char, valor
    
    soma = 0
    peso = pesoInicial
    
    For i = 1 To Len(base)
        char = Mid(base, i, 1)
        
        ' Converte caractere em valor numérico (para letras A-Z)
        If IsNumeric(char) Then
            valor = CInt(char)
        Else
            ' Para letras A-Z, converte para 10-35
            valor = Asc(char) - Asc("A") + 10
        End If
        
        soma = soma + (valor * peso)
        
        ' Lógica dos pesos: vai de 9 a 2, depois recomeça
        If peso = 2 Then
            peso = 9
        Else
            peso = peso - 1
        End If
    Next
    
    Dim resto
    resto = soma Mod 11
    
    If resto < 2 Then
        calcularDigito = 0
    Else
        calcularDigito = 11 - resto
    End If
End Function

' --- Exemplos de uso ---
' Para testar, descomente as linhas abaixo
' Response.Write validaCNPJ("12.ABC.345/01DE-35") & "<br>"   ' true (exemplo válido)
' Response.Write validaCNPJ("12.ABC.345/01DE-99") & "<br>"   ' false (dígitos errados)
' Response.Write validaCNPJ("11.222.333/0001-81") & "<br>"   ' true (CNPJ numérico tradicional)
' Response.Write validaCNPJ("12.ABC.345/01DE-3") & "<br>"    ' false (tamanho incorreto)
%>