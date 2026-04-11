Sub IntegrarComWord()

    On Error GoTo TratarErro
    
    Dim wordApp As Object
    Dim doc As Object
    
    Set wordApp = CriarInstanciaWord()
    Set doc = AbrirDocumento(wordApp, ObterCaminhoArquivo())
    
    SubstituirPlaceholders doc, "Ocorrencia"

    SalvarComoPDF doc, Worksheets("Ocorrencia").Range("C2").Value
    
    Exit Sub

TratarErro:
    MsgBox "Erro ao gerar documento: " & Err.Description, vbCritical

End Sub


'========================
' CRIA INSTÂNCIA DO WORD
'========================
Private Function CriarInstanciaWord() As Object
    Dim app As Object
    Set app = CreateObject("Word.Application")
    app.Visible = True
    Set CriarInstanciaWord = app
End Function


'========================
' ABRE DOCUMENTO
'========================
Private Function AbrirDocumento(wordApp As Object, caminho As String) As Object
    Set AbrirDocumento = wordApp.Documents.Open(caminho)
End Function


'========================
' CAMINHO DO ARQUIVO
'========================
Private Function ObterCaminhoArquivo() As String
    ObterCaminhoArquivo = "C:\Users\Marcello Barros\Documents\marcello barros\11 - projetos-caml\VBA - Relatórios Gerenciais\Teste.docx"
End Function


'========================
' SUBSTITUI VALORES
'========================
Private Sub SubstituirPlaceholders(doc As Object, nomePlanilha As String)

    Dim mapa As Object
Set mapa = CriarMapaSubstituicoes()

Dim chave As Variant
Dim conteudo As Object
Dim ws As Worksheet
Dim valor As String

Set ws = Worksheets(nomePlanilha)
Set conteudo = doc.Content

For Each chave In mapa.Keys
    
    valor = ""
    
    If Not IsError(ws.Range(mapa(chave)).Value) Then
        valor = CStr(ws.Range(mapa(chave)).Value)
    End If
    
    ExecutarSubstituicao conteudo, CStr(chave), valor
    
Next chave

End Sub


'========================
' MAPA DE CAMPOS
'========================
Private Function CriarMapaSubstituicoes() As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "Sequencia", "A2"
    dict.Add "Num_Socio", "B2"
    dict.Add "Nome_Socio", "C2"
    dict.Add "Tipo_Ocorrencia", "D2"
    dict.Add "Data_Ocorrencia", "E2"
    dict.Add "Desc_Ocorrencia", "F2"
    
    Set CriarMapaSubstituicoes = dict

End Function


'========================
' EXECUTA SUBSTITUIÇÃO
'========================
Private Sub ExecutarSubstituicao(conteudo As Object, textoBusca As String, ByVal textoSubstituto As String)

    Const wdReplaceAll = 2
    Const wdFindContinue = 1

    With conteudo.Find
        .Text = textoBusca
        .Replacement.Text = textoSubstituto
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With

End Sub

Private Sub SalvarComoPDF(doc As Object, identificador As String)

    Dim caminho As String
    
    caminho = "C:\Users\Marcello Barros\Documents\marcello barros\11 - projetos-caml\VBA - Relatórios Gerenciais\Historico Relatórios\Relatorio_" & identificador & ".pdf"
    
    doc.ExportAsFixedFormat _
        OutputFileName:=caminho, _
        ExportFormat:=17

End Sub