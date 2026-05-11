Sub ImgConvert()
    Dim conn As Object
    Dim rs As Object
    Dim stream As Object

    Set conn = DbConect()
    If conn Is Nothing Then
        MsgBox "Falha na conexão com o banco.", vbCritical
        Exit Sub
    End If

    Set rs = SelectColumn(conn)
    If rs Is Nothing Then
        MsgBox "Falha ao buscar dados.", vbCritical
        conn.Close
        Exit Sub
    End If

    Dim linha As Integer
    linha = 1

    Do While Not rs.EOF
        Call ConvertImg(rs)
        rs.MoveNext
    Loop

    rs.Close
    conn.Close

    MsgBox "Imagens importadas com sucesso!"
End Sub

Private Function DbConect() As Object
    Dim conn As Object
    Dim strConn As String

    strConn = "Provider=xxxxxx;" & _
              "Data Source=xxxxxx;" & _
              "Initial Catalog=xxxxxx;" & _
              "User ID=xxxxxx;" & _
              "Password=xxxxxx;"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn

    Set DbConect = conn
End Function

Private Function SelectColumn(conn As Object) As Object
    Dim rs As Object
    Dim strSearch As String
    Dim numSocio As Long

    numSocio = Worksheets("pesquisa").Range("B5").Value
    strSearch = "SELECT CLB_SocioFoto FROM CLB_SOCIO WHERE CLB_SocioID = " & numSocio

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSearch, conn

    Set SelectColumn = rs
End Function

Private Sub ConvertImg(rs As Object)
    Dim stream As Object
    Dim caminho As String

    caminho = "M:\ADM_FIN\GER_FIN\14 - Diversos Marcello\15-sugestoes-reclamacoes\fotos-arquivadas\foto.jpg"

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write rs("CLB_SocioFoto").Value
    stream.SaveToFile caminho, 2
    stream.Close

    Dim ws As Worksheet
    Dim cell As Range
    Set ws = Worksheets("pesquisa")
    Set cell = ws.Range("Q5")

    Dim pic As Object
    For Each pic In ws.Pictures
        If pic.Left >= cell.Left And pic.Top >= cell.Top Then
            pic.Delete
            Exit For
        End If
    Next pic

    Dim novaPic As Object
    Set novaPic = ws.Pictures.Insert(caminho)
    With novaPic
        .Left = cell.Left
        .Top = cell.Top
        .Width = cell.Width
        .Height = cell.Height
    End With
End Sub
