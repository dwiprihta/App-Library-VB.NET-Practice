Imports MySql.Data.MySqlClient
Module Koneksi
    Public conn As MySqlConnection
    Public dr As MySqlDataReader
    Public da As MySqlDataAdapter
    Public cmd As MySqlCommand
    Public ds As DataSet
    Public simpan, edit, hapus As String

    Public Sub bukadb()
        Dim sqlconn As String
        sqlconn = "server=localhost;uid=root;password=;database=perpus"
        conn = New MySqlConnection(sqlconn)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
    End Sub
End Module
