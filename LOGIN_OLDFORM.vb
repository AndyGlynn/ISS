Public Class LOGIN_OLD
    Private DefaultUserName As String = "Admin"
    Private DefaultPwd As String = "admin"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Application.Exit()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim log_In As New USER_LOGIC
        log_In.Login_To_System()
        Me.Close()
        Main.TSMain.Enabled = True

    End Sub

    Private Sub LOGIN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtPWD.Text = DefaultPwd
        Me.txtUser.Text = DefaultUserName

    End Sub
End Class
