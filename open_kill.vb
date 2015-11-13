Public Class open_kill

    Public Sub New()
        Kill.Contact1 = ConfirmingSingleRecord.txtContact1.Text
        Kill.Contact2 = ConfirmingSingleRecord.txtContact2.Text
        Kill.frm = "ConfirmingSingleRecord"
        Kill.ID = ConfirmingSingleRecord.ID

        Kill.ShowInTaskbar = False
        Kill.ShowDialog()
    End Sub


End Class
