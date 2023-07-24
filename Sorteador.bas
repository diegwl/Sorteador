Attribute VB_Name = "M�dulo1"
Sub GerarNumeros()
Dim i                           As Integer
Dim j                           As Integer
Dim k                           As Integer
Dim l                           As Integer
Dim bRandomOk                   As Boolean
Dim valor_aleatorio             As Integer
Dim valor_maior                 As Integer
Dim total_numeros_gerados       As Integer
Dim total_numeros_para_gerar    As Integer
Dim iControleGerar              As Integer
Dim iColunaCelula               As Integer
Dim cont                        As Integer
Dim cont_nome                   As Integer


    Sheets("Dados").Range("C3:C100").ClearContents
    Sheets("Tela de Sorteio").Range("B5:G16").ClearContents
    
    valor_maior = 6    'Informe o maior n�mero que poder� ser gerado
    total_numeros_para_gerar = 18      'Informe a quantidade de n�meros aleat�rios que deseja gerar
    total_numeros_gerados = 0
    iLinhaCelulaInicial = 3     'Informe a linha da primeira c�lula que ser� escrito o n�mero
    iColunaCelula = 3   'Informe a coluna. Exemplo: Coluna B = 2
    iControleGerar = total_numeros_para_gerar + iLinhaCelulaInicial - 1

    'Gera quantos n�meros forem indicados na vari�vel 'total_numeros_gerados'
    For i = iLinhaCelulaInicial To iControleGerar
        total_numeros_gerados = total_numeros_gerados + 1

        'Fica executando a gera��o de um novo n�mero enquanto houver duplicidade
        Do

            'Gera um novo n�mero
            Randomize   'Sempre utilize esta fun��o antes de chamar Rnd
            valor_aleatorio = Int((valor_maior * Rnd) + 1)
            bRandomOk = True

            'Verifica se j� saiu este n�mero
            cont = 0
            For j = iLinhaCelulaInicial To i
                If Sheets("Dados").Cells(j, iColunaCelula).Value = valor_aleatorio Then
                    cont = cont + 1
                End If
            If cont >= 3 Then
                bRandomOk = False
            End If
            Next j

        Loop While bRandomOk = False

        'Escreve o n�mero na c�lula
        Sheets("Dados").Cells(i, iColunaCelula).Value = valor_aleatorio
        
    Next i
    
    For k = 0 To valor_maior
        cont_nome = 0
        For l = iLinhaCelulaInicial To i
                If Sheets("Dados").Cells(l, iColunaCelula).Value = k Then
                    cont_nome = cont_nome + 1
                    Sheets("Tela de Sorteio").Cells(cont_nome + 4, k + 1).Value = Sheets("Dados").Cells(l, 1).Value
                End If
        Next l
    Next k
        

    MsgBox "Nomes Sorteados", vbInformation

End Sub

Sub Clear()
    Sheets("Dados").Range("C3:C100").ClearContents
    Sheets("Tela de Sorteio").Range("B5:G16").ClearContents
    
    MsgBox "Campos Limpos", vbInformation
End Sub
