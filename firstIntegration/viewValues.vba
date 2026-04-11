Function SomaDoisNumeros(valor1 As Double, valor2 As Double) As Double
  SomaDoisNumeros = valor1 + valor2
End Function

Function NumNegativo(valor1 As Double) As Boolean
  If valor1 < 0 Then
    NumNegativo = True
  Else
    NumNegativo = False
  End If
  
End Function

Sub ExibirValor()
  Dim resultado As Double
  Dim vNegativo As Boolean

  resultado = SomaDoisNumeros(Worksheets("Plan1").Range("A1").Value, Worksheets("Plan1").Range("B1").Value)

  vNegativo = NumNegativo(resultado)

  MsgBox "A soma dos números é: " & resultado & vbNewLine & "O resultado é negativo? " & vNegativo & IIf(numNegativo, "Sim", "Não")

End Sub