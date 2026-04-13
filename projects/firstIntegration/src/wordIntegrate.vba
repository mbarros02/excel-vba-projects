Sub IntegrarComWord()

    On Error GoTo TratarErro
    
    Dim wordApp As Object
    Dim doc As Object
    
    Set wordApp = CriarInstanciaWord()
    Set doc = AbrirDocumento(wordApp, ObterCaminhoArquivo())
    
    SubstituirPlaceholders doc, "pesquisa"

    SalvarComoPDF doc, Worksheets("pesquisa").Range("F4").Value

    Set doc = FecharDocumento(wordApp, ObterCaminhoArquivo())

    MsgBox "Relatório Finalizado com Sucesso!" & vbNewLine & "Solicitação - " & Worksheets("pesquisa").Range("F4").Value 
    
    Exit Sub

TratarErro:
    MsgBox "Erro ao gerar documento: " & Err.Description, vbCritical

End Sub


Private Function CriarInstanciaWord() As Object
    Dim app As Object
    Set app = CreateObject("Word.Application")
    app.Visible = False
    Set CriarInstanciaWord = app
End Function


Private Function AbrirDocumento(wordApp As Object, caminho As String) As Object
    Set AbrirDocumento = wordApp.Documents.Open(caminho)
End Function

Private Function FecharDocumento(wordApp As Object, caminho As String) As Object
    Set FecharDocumento = wordApp.Documents.Close(caminho)
End Function

Private Function ObterCaminhoArquivo() As String
    ObterCaminhoArquivo = "M:\ADM_FIN\GER_FIN\14 - Diversos Marcello\01-relatorios\08 - relatorios-macros\relatorio-solicitacoes\template_solicitacoes.docx"
End Function


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
        
        ' Substitui diretamente no Bookmark
        If doc.Bookmarks.Exists(chave) Then
            doc.Bookmarks(chave).Range.Text = valor
        End If
    Next chave

End Sub


Private Function CriarMapaSubstituicoes() As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "num_solicitacao", "A4"
    dict.Add "num_socio", "E4"
    dict.Add "nome_socio", "F4"
    dict.Add "email_socio", "G4"
    dict.Add "assunto_solicitacao", "N4"
    dict.Add "tipo_solicitacao", "M4"
    dict.Add "data_solicitacao", "AR4"
    dict.Add "texto_solicitacao", "AK4"
    
    Set CriarMapaSubstituicoes = dict

End Function


Private Sub SalvarComoPDF(doc As Object, identificador As String)

    Dim caminho As String
    
    caminho = "M:\ADM_FIN\GER_FIN\14 - Diversos Marcello\15-sugestoes-reclamacoes\SOLICITAÇÃO_" & identificador & ".pdf"
    
    doc.ExportAsFixedFormat _
        OutputFileName:=caminho, _
        ExportFormat:=17

End Sub