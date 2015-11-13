Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class Alert_Logic_Tick
    Private cntOfAlertsTick As Integer = 0
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public Sub New(ByVal UserName As String)
        Dim cnt As Integer = CountAlertsTick(UserName)
        Me.CountOfAlertsTick = cnt
    End Sub

    Public Property CountOfAlertsTick() As Integer
        Get
            Return cntOfAlertsTick
        End Get
        Set(ByVal value As Integer)
            cntOfAlertsTick = value
        End Set
    End Property

    Public Function CountAlertsTick(ByVal UserName As String)
        Try
            Dim cmdCNT As SqlCommand = New SqlCommand("select count (id) from AlertTable where Username = @USR and (select ExecutionDate + AlertTime ) = substring ({fn current_timestamp ()},0,17)", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@USR", UserName)
            cmdCNT.Parameters.Add(param1)
            cmdCNT.CommandType = CommandType.Text
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdCNT.ExecuteReader
            While r1.Read
                CountOfAlertsTick = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
            Return CountOfAlertsTick
        Catch ex As Exception
            cnn.Close()
            Return CountOfAlertsTick
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ALERT_LOGIC", "UserName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "CountAlertsTick")
        End Try
    End Function
End Class
