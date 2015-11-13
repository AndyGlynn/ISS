Public Class frmErrorLogs

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim x As New ErrorLogFlatFile
        x.ClearLogs("Both")
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '' instantiate object
        '' 
        Dim err As New ErrorLogFlatFile

        Dim str As String = ""
        str = Me.ComboBox1.Text
        If str.ToString.Length <= 0 Then
            Exit Sub
        ElseIf str.ToString.Length >= 1 Then
            Select Case str
                Case Is = "Network - Client"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "Network", "Test")
                    Exit Select
                Case Is = "Network - Server"
                    err.WriteLog("Test", "test", "system exception", "Server", "ADMIN", "Network", "Test")
                    Exit Select
                Case Is = "SQL - Client"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "SQL", "Test")
                    Exit Select
                Case Is = "SQL - Server"
                    err.WriteLog("Test", "test", "system exception", "Server", "ADMIN", "SQL", "Test")
                    Exit Select
                Case Is = "FrontEnd - Client ONLY"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "Front_End", "Test")
                    Exit Select
                Case Is = "TAPI - CLient ONLY"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "TAPI", "Test")
                    Exit Select
                Case Is = "ReportingSVCS - Client"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "ReportingSVCS", "Test")
                    Exit Select
                Case Is = "ReportingSVCS - Server"
                    err.WriteLog("Test", "test", "system exception", "Server", "ADMIN", "ReportingSVCS", "Test")
                    Exit Select
                Case Is = "Mappoint - Client ONLY"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "Mappoint", "Test")
                    Exit Select
                Case Is = "FileIO - Client ONLY"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "File_IO", "Test")
                    Exit Select
                Case Is = "Chat - Server ONLY"
                    err.WriteLog("Test", "test", "System exception", "Server", "ADMIN", "Chat", "Test")
                    Exit Select
                Case Is = "XFer - Client"
                    err.WriteLog("Test", "test", "system exception", "Client", "ADMIN", "XFER", "Test")
                    Exit Select
                Case Is = "XFer - Server"
                    err.WriteLog("Test", "test", "system exception", "Server", "ADMIN", "XFER", "Test")
                    Exit Select
                Case Is = "General_Client - Server ONLY"
                    err.WriteLog("Test", "test", "system exception", "Server", "ADMIN", "General_Client", "Test")
                    Exit Select
            End Select
        End If

    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        Dim err As New ErrorLogFlatFile

        Dim str As String = Me.ComboBox2.Text
        If str.ToString.Length <= 0 Then
            Exit Sub
        ElseIf str.ToString.Length >= 1 Then
            err.ClearLogs(str)
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedValueChanged
        Dim err As New ErrorLogFlatFile

        Dim str As String = Me.ComboBox3.Text
        If str.ToString.Length <= 0 Then
            Exit Sub
        ElseIf str.ToString.Length >= 1 Then
            err.Get_Logs(str)
        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

   
    Private Sub frmErrorLogs_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
