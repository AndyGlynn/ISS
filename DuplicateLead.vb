Public Class DuplicateLead
    Public CloseMethod As String = ""
    Public MapPointVerified As Boolean
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.lstDupes.Items.Clear()
        Me.CloseMethod = ""
        Me.Close()
    End Sub

    Private Sub DuplicateLead_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.lstDupes.Items.Clear()
        Select Case CloseMethod
            Case "SAVE"
                EnterLead.Reset()
                EnterLead.Close()
                Exit Select
            Case "SAVEANDNEW"
                EnterLead.Reset()
                Exit Select
            Case ""
                Exit Select
        End Select
    End Sub

    
    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim c As New ENTER_LEAD.InsertEnterLead
        c.InsertLead(MapPointVerified)
        Me.lstDupes.Items.Clear()
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub DuplicateLead_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
