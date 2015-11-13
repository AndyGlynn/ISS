Public Class frmBackup

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim b As New DBASE_Backup
        b.Backup()
        b.Get_Existing()
        'System.Diagnostics.Process.Start("\\ekg1\iss\backups")

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim b As New DBASE_Backup
        b.Get_Existing()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.ListView1.Items.Clear()
        Me.Close()
    End Sub
End Class
