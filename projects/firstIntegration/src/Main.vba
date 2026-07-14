Sub GerarRelatorio()
    On Error GoTo TratarErro

    Dim caminhoFoto As String
    caminhoFoto = ImageService.BuscarEConverterFoto()
    If caminhoFoto = "" Then
        MsgBox "Foto do sócio não encontrada.", vbExclamation
        Exit Sub
    End If

    ImageService.ExibirNaPlanilha caminhoFoto

    WordService.GerarRelatorio caminhoFoto

    MsgBox "Relatório gerado com sucesso!" & vbNewLine & _
           "Solicitação: " & Worksheets("pesquisa").Range("H5").Value & "_" & Worksheets("pesquisa").Range("A5").Value

    Exit Sub

TratarErro:
    MsgBox "Erro em GerarRelatorio: " & Err.Description, vbCritical
End Sub

