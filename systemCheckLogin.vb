Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient


Public Class systemCheckLogin
    Public cntOfRecords As Integer = 0
    Public Const sqlCNX As String = "Server=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1"

    Public Function CntRecs()
        Dim cnt_cnx As SqlConnection = New SqlConnection(sqlCNX)
        Dim cmdCNT As SqlCommand = New SqlCommand("SELECT count(ID) from EnterLead;", cnt_cnx)
        Dim cnt As Integer = 0
        cnt_cnx.Open()
        cnt = cmdCNT.ExecuteScalar
        cnt_cnx.Close()
        cnt_cnx = Nothing
        Return cnt
    End Function

    Public Function CheckNetworkStatus()
        Dim status As Boolean
        If My.Computer.Network.IsAvailable = True Then
            status = True
        ElseIf My.Computer.Network.IsAvailable = False Then
            status = False
        End If
        Return status
    End Function

    Public Function CheckSQLStatus()
        Dim status As Boolean
        Dim test_cnx As SqlConnection = New SqlConnection(sqlCNX)

        Try
            test_cnx.Open()
            If test_cnx.State = ConnectionState.Open Then
                status = True
            End If
        Catch ex As Exception
            If test_cnx.State = ConnectionState.Broken Or test_cnx.State = ConnectionState.Closed Then
                status = False
            End If
        End Try
        Return status
    End Function

End Class
