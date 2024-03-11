Imports System.Data.OleDb
Imports System.Linq
Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml.XPath
Imports System.Data
Imports System.Xml

Module mdlfunction


    Public OleDbCommand As New OleDbCommand
    Public OleDbDataAdapter As New OleDbDataAdapter

    Public Connect As New OleDbConnection

    Public Function ConnectionString_MDB() As String
        ConnectionString_MDB = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                  "Data Source= " & Application.StartupPath & "\Gaffer_DB.mdb"

    End Function
    Public Sub execute(ByVal sqlstring As String)
        If Connect.State = ConnectionState.Open Then
            Connect.Close()
        End If
        Connect.Open()
        OleDbCommand = New OleDbCommand(sqlstring, Connect)
        OleDbCommand.ExecuteNonQuery()
        OleDbCommand.Dispose()
        OleDbCommand = Nothing
    End Sub
    Public Function GET_RECORDS(ByVal sqlstring As String, ByVal connect As OleDbConnection) As DataTable
        If connect.State = ConnectionState.Open Then
            connect.Close()
        End If
        connect.Open()
        OleDbCommand = New OleDbCommand(sqlstring, connect)
        OleDbDataAdapter.SelectCommand = OleDbCommand
        Dim Temptable As New DataTable
        OleDbDataAdapter.Fill(Temptable)
        OleDbCommand.Dispose()
        OleDbCommand = Nothing
        connect.Close()
        Return Temptable
    End Function

End Module
