Option Explicit

Private Const CONN_STRING As String = _
    "Provider=xxxxx;" & _
    "Data Source=xxxxxxx;" & _
    "Initial Catalog=xxxxxx;" & _
    "User ID=xxxxxx;" & _
    "Password=xxxxxx;"

Public Function AbrirConexao() As Object
    On Error GoTo TratarErro

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open CONN_STRING
    Set AbrirConexao = conn
    Exit Function

TratarErro:
    MsgBox "Erro ao conectar ao banco: " & Err.Description, vbCritical
    Set AbrirConexao = Nothing
End Function

Public Function BuscarFotoSocio(conn As Object) As Object
    On Error GoTo TratarErro

    Dim numSocio As Long
    numSocio = Worksheets("pesquisa").Range("B5").Value

    Dim sql As String
    sql = "SELECT xxxxxxx FROM xxxxxxx WHERE xxxxxxx = " & numSocio

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    Set BuscarFotoSocio = rs
    Exit Function

TratarErro:
    MsgBox "Erro ao buscar foto: " & Err.Description, vbCritical
    Set BuscarFotoSocio = Nothing
End Function

