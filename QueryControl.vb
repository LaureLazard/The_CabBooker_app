Imports System.Data.OleDb
Public Class QueryControl
    'CREAT DB CONNECTION
    Private DBCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=CabBooker.DB.accdb")

    'PREPARE COMMAND CALLS
    Private DbCmd As OleDbCommand

    'PREPARE DATA COLLECTOR
    Public DBDA As OleDbDataAdapter
    Public DBDT As DataTable

    'QUERY PARAMETERS
    Public Params As New List(Of OleDbParameter)

    'QUERY STATISTICS
    Public RecCount As Integer
    Public Exception As String

    Public Sub ExecQuery(Query As String)
        'RESET QUERY STATS
        RecCount = 0
        Exception = ""
        Try
            'OPEN A CONNECTION
            DBCon.Open()

            'CREATE DB COMMAND
            DbCmd = New OleDbCommand(Query, DBCon)

            'LOAD PARAMETERS INTO COMMAND
            Params.ForEach(Sub(Par) DbCmd.Parameters.Add(Par))

            'CLEAR PARAMETERS LIST
            Params.Clear()

            'EXECUTE COMMAND & FILL DATA
            DBDT = New DataTable
            DBDA = New OleDbDataAdapter(DbCmd)
            RecCount = DBDA.Fill(DBDT)
        Catch ex As Exception
            Exception = ex.Message
        End Try

        'CLOSE CONNECTION
        If DBCon.State = ConnectionState.Open Then DBCon.Close()
    End Sub

    'INCLUDE QUERY AND COMMAND PARAMETERS
    Public Sub AddP(Name As String, Value As Object)
        Dim NewParams As New OleDbParameter(Name, Value)
        Params.Add(NewParams)
    End Sub

End Class
