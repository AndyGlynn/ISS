Public Class frmViewEditPrimaryLeadSources

    Private Sub frmViewEditPrimaryLeadSources_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim y As New ViewEditPrimaryLeadSources
        y.PopulateList()
        Dim d
        For d = 0 To y.ArPLS.Count - 1
            Me.chkPrimaryLeadSources.Items.Add(y.ArPLS(d).ToString)
        Next

    End Sub

    Private Sub ResetDefault()
        Me.chkPrimaryLeadSources.Items.Clear()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim str As String = ""
        Dim y As New ViewEditPrimaryLeadSources

        Try
            str = InputBox$("Enter a new Primary Lead Source.", "New Primary Lead Source")
            If str.ToString.Length <= 0 Then
                MsgBox("You cannot add a blank Primary Lead Source.", MsgBoxStyle.Critical, "Error Adding Primary Lead Source")
                Exit Sub
            ElseIf str.ToString.Length >= 1 Then
                y.Add_PLS(str)
                Me.chkPrimaryLeadSources.Items.Clear()
                y.PopulateList()

                Dim d
                For d = 0 To y.ArPLS.Count - 1
                    Me.chkPrimaryLeadSources.Items.Add(y.ArPLS(d).ToString)
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ResetDefault()
        Me.Close()
    End Sub
End Class
