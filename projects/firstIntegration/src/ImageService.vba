Option Explicit

Private Const CAMINHO_FOTO As String = _
    "xxxxxxx.jpg"

Private Const CELULA_FOTO As String = "Q5"
Private Const PLANILHA_PESQUISA As String = "pesquisa"

Public Function BuscarEConverterFoto() As String
    Dim conn As Object
    Dim rs As Object

    Set conn = DatabaseService.AbrirConexao()
    If conn Is Nothing Then Exit Function

    Set rs = DatabaseService.BuscarFotoSocio(conn)
    If rs Is Nothing Or rs.EOF Then
        rs.Close
        conn.Close
        Exit Function
    End If

    SalvarBlob rs("xxxxxx").Value, CAMINHO_FOTO

    rs.Close
    conn.Close

    BuscarEConverterFoto = CAMINHO_FOTO
End Function

Private Sub SalvarBlob(blob As Variant, caminho As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write blob
    stream.SaveToFile caminho, 2
    stream.Close
End Sub

Public Sub ExibirNaPlanilha(caminho As String)
    Dim ws As Worksheet
    Dim cell As Range
    Set ws = Worksheets(PLANILHA_PESQUISA)
    Set cell = ws.Range(CELULA_FOTO)

    LimparImagemNaCelula ws, cell

    Dim pic As Object
    Set pic = ws.Pictures.Insert(caminho)
    With pic
        .Left = cell.Left
        .Top = cell.Top
        .Width = cell.Width
        .Height = cell.Height
    End With
End Sub

Private Sub LimparImagemNaCelula(ws As Worksheet, cell As Range)
    Dim pic As Object
    For Each pic In ws.Pictures
        If pic.Left >= cell.Left And pic.Top >= cell.Top _
        And pic.Left < cell.Left + cell.Width _
        And pic.Top < cell.Top + cell.Height Then
            pic.Delete
            Exit For
        End If
    Next pic
End Sub

