Option Explicit

Private Const CAMINHO_TEMPLATE As String = _
    "xxxxxxxxx.docx"

Private Const CAMINHO_PDF As String = _
    "xxxxxxxxxxx"

Private Const PLANILHA_PESQUISA As String = "pesquisa"

Public Sub GerarRelatorio(caminhoFoto As String)
    On Error GoTo TratarErro

    Dim wordApp As Object
    Dim doc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    Set doc = wordApp.Documents.Open(CAMINHO_TEMPLATE)

    SubstituirTextos doc
    SubstituirFoto doc, caminhoFoto
    
    Dim solicitacao As String
    Dim identificador As String
    solicitacao = Worksheets(PLANILHA_PESQUISA).Range("A5").Value
    identificador = Worksheets(PLANILHA_PESQUISA).Range("H5").Value
    ExportarPDF doc, identificador, solicitacao

    doc.Close False
    wordApp.Quit
    Exit Sub

TratarErro:
    If Not doc Is Nothing Then doc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    MsgBox "Erro ao gerar Word/PDF: " & Err.Description, vbCritical
End Sub

Private Sub SubstituirTextos(doc As Object)
    Dim ws As Worksheet
    Set ws = Worksheets(PLANILHA_PESQUISA)

    Dim mapa As Object
    Set mapa = MapaBookmarks()

    Dim chave As Variant
    Dim valor As String

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

Private Sub SubstituirFoto(doc As Object, caminhoImagem As String)
    If caminhoImagem = "" Then Exit Sub
    If Not doc.Bookmarks.Exists("foto_socio") Then Exit Sub

    Dim rng As Object
    Set rng = doc.Bookmarks("foto_socio").Range

    Dim posLeft As Single
    Dim posTop As Single
    posLeft = rng.Information(1)
    posTop = rng.Information(2)
    rng.Text = ""

    Dim shape As Object
    Set shape = doc.Shapes.AddPicture( _
        Filename:=caminhoImagem, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=posLeft, _
        Top:=posTop, _
        Width:=100, _
        Height:=100, _
        Anchor:=rng)

    With shape
        .LockAspectRatio = True
        .WrapFormat.Type = 3
        .WrapFormat.Side = 0
    End With
End Sub

Private Sub ExportarPDF(doc As Object, identificador As String, solicitacao As String)
    Dim caminho As String
    caminho = CAMINHO_PDF & "SOLICITAÇÃO_" & identificador & "_" & solicitacao & ".pdf"
    doc.ExportAsFixedFormat OutputFileName:=caminho, ExportFormat:=17
End Sub

Private Function MapaBookmarks() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "data_solicitacao", "C5"
    dict.Add "num_solicitacao", "A5"
    dict.Add "nome_socio", "H5"
    dict.Add "num_socio", "G5"
    dict.Add "celular_socio", "J5"
    dict.Add "email_socio", "I5"
    dict.Add "assunto_solicitacao", "L5"
    dict.Add "tipo_solicitacao", "L5"
    dict.Add "status", "F5"
    dict.Add "texto_solicitacao", "K5"
    Set MapaBookmarks = dict
End Function

