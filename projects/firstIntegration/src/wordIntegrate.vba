Sub IntegrarComWord()

    On Error GoTo TratarErro
    
    Dim wordApp As Object
    Dim doc As Object
    
    Set wordApp = CriarInstanciaWord()
    Set doc = AbrirDocumento(wordApp, ObterCaminhoArquivo())
    
    SubstituirPlaceholders doc, "pesquisa"

    Dim caminhoFoto As String
    caminhoFoto = "M:\ADM_FIN\GER_FIN\14 - Diversos Marcello\15-sugestoes-reclamacoes\fotos-arquivadas\foto.jpg"
    SubstituirImagemBookmark doc, "foto_socio", caminhoFoto

    SalvarComoPDF doc, Worksheets("pesquisa").Range("H5").Value

    doc.Close False
    wordApp.Quit

    MsgBox "Relatório Finalizado com Sucesso!" & vbNewLine & "Solicitação - " & Worksheets("pesquisa").Range("H5").Value
    
    Exit Sub

TratarErro:
    doc.Close False
    wordApp.Quit
    MsgBox "Erro ao gerar documento: " & Err.Description, vbCritical

End Sub


Private Function CriarInstanciaWord() As Object
    Dim app As Object
    Set app = CreateObject("Word.Application")
    app.Visible = True
    Set CriarInstanciaWord = app
End Function


Private Function AbrirDocumento(wordApp As Object, caminho As String) As Object
    Set AbrirDocumento = wordApp.Documents.Open(caminho)
End Function

Private Function ObterCaminhoArquivo() As String
    ObterCaminhoArquivo = "M:\ADM_FIN\GER_FIN\14 - Diversos Marcello\01-relatorios\08 - relatorios-macros\relatorio-solicitacoes\template_solicitacoes.docx"
End Function

Private Sub SubstituirImagemBookmark(doc As Object, nomeBookmark As String, caminhoImagem As String)
    If caminhoImagem = "" Then Exit Sub
    If Not doc.Bookmarks.Exists(nomeBookmark) Then Exit Sub

    Dim rng As Object
    Set rng = doc.Bookmarks(nomeBookmark).Range

    ' Limpa o texto do bookmark e insere a imagem inline
    rng.Text = ""
    Dim shape As Object
    Set shape = doc.InlineShapes.AddPicture( _
        FileName:=caminhoImagem, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=rng)

    ' Ajusta o tamanho se necessário (opcional)
    shape.LockAspectRatio = True
    shape.Height = 100 ' altura em pontos (~3,5 cm) — ajuste conforme o template
End Sub


Private Sub SubstituirPlaceholders(doc As Object, nomePlanilha As String)

    Dim mapa As Object
    Set mapa = CriarMapaSubstituicoes()

    Dim chave As Variant
    Dim ws As Worksheet
    Dim valor As String

    Set ws = Worksheets(nomePlanilha)

    For Each chave In mapa.Keys
        valor = ""
        If Not IsError(ws.Range(mapa(chave)).Value) Then
            valor = CStr(ws.Range(mapa(chave)).Value)
        End If
        
        If doc.Bookmarks.Exists(chave) Then
            doc.Bookmarks(chave).Range.Text = valor
        End If
    Next chave

End Sub

Private Function CriarMapaSubstituicoes() As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    dict.Add "data_relatorio", "B1"
    dict.Add "celular_socio", "J5"
    dict.Add "num_solicitacao", "A5"
    dict.Add "num_socio", "G5"
    dict.Add "nome_socio", "H5"
    dict.Add "email_socio", "I5"
    dict.Add "assunto_solicitacao", "L5"
    dict.Add "tipo_solicitacao", "L5"
    dict.Add "data_solicitacao", "C5"
    dict.Add "texto_solicitacao", "K5"
    
    Set CriarMapaSubstituicoes = dict

End Function



Private Function SalvarComoPDF(doc As Object, identificador As String) As String

    Dim caminho As String
    
    caminho = "M:\ADM_FIN\GER_FIN\14 - Diversos Marcello\15-sugestoes-reclamacoes\SOLICITAÇÃO_" & identificador & ".pdf"
    
    doc.ExportAsFixedFormat _
        OutputFileName:=caminho, _
        ExportFormat:=17

    SalvarComoPDF = caminho

End Function

End Sub