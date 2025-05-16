Sub calculator()
' Isto é um teste

    Dim valor As Integer
    Dim texto As String
    Dim resultado As Integer
    
    Dim folha As Worksheet
    
    Set folha = Worksheets("Folha1")
    valor = 1000
    texto = "+"
    resultado = valor + 4
    
    Range("A1").Value = valor
    
    valor = folha.Range("D1").Value
    
    folha.Range("A2").Value = texto
    
    Range("A3").Value = resultado
    
    MsgBox "apresenta o valor de D1 = " & valor
    
    folha.Range("A4").FormulaLocal = "=MAIÚSCULAS(A2)"
    
    For i = 1 To 10
        folha.Cells(t, i) = i
    Next i
    
    
    
End Sub
