
Randomize
num1 = Int((100 - 1 + 1) * Rnd + 1)
MsgBox "Primeiro número sorteado: " & num1

num2 = Int((100 - 1 + 1) * Rnd + 1)
MsgBox "Segundo número sorteado: " & num2
                             
operacao = Int((4 - 1 + 1) * Rnd + 1)

Select Case operacao
    Case 1
        resultado = num1 + num2
        operacaoSimbolo = "+"
    Case 2
        resultado = num1 - num2
        operacaoSimbolo = "-"
    Case 3
        resultado = num1 * num2
        operacaoSimbolo = "*"
    Case 4
        If num2 <> 0 Then
            resultado = num1 / num2
            operacaoSimbolo = "/"
        Else
            resultado = "Indefinido (divisão por zero)"
        End If
End Select


Select Case operacao
    Case 1
        MsgBox "Operação sorteada: Adição (+)"
    Case 2
        MsgBox "Operação sorteada: Subtração (-)"
    Case 3
        MsgBox "Operação sorteada: Multiplicação (*)"
    Case 4
        MsgBox "Operação sorteada: Divisão (/)"
End Select


If operacao = 4 And num2 = 0 Then
    MsgBox "Números sorteados: " & num1 & " e " & num2 & vbCrLf & _
           "Operação: " & operacaoSimbolo & vbCrLf & _
           "Resultado: " & resultado
Else
    MsgBox "Números sorteados: " & num1 & " e " & num2 & vbCrLf & _
           "Operação: " & num1 & " " & operacaoSimbolo & " " & num2 & vbCrLf & _
           "Resultado: " & resultado
End If