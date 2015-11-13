Public Class openset

    Private Sub openset_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Visible = False
        Dim x As New opensetappt(Sales, Sales.ID, Sales.txtApptDate.Text, Sales.txtApptTime.Text, Sales.txtContact1.Text, Sales.txtContact2.Text)
        Me.Close()
    End Sub
End Class
