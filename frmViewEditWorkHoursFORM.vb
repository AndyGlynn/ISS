Public Class frmViewEditWorkHours

    
    Private Sub ResetDefault()
        Me.chklstWorkHours.Items.Clear()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            Dim y As New ViewEditWorkHours
            Dim t As Object
            Dim str As String = ""
            For Each t In Me.chklstWorkHours.CheckedItems
                str = t.ToString
                If str.ToString.Length <= 0 Then
                    '' do nothing
                ElseIf str.ToString.Length >= 1 Then
                    y.DeleteWorkHour(str)
                End If
            Next
            ResetDefault()
            y.PopulateWorkHours()
            Dim d
            For d = 0 To y.ArWorkHours.Count - 1
                Me.chklstWorkHours.Items.Add(y.ArWorkHours(d), False)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmViewEditWorkHours_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ResetDefault()
        Dim x As New ViewEditWorkHours
        x.PopulateWorkHours()
        Dim d
        For d = 0 To x.ArWorkHours.Count - 1
            Me.chklstWorkHours.Items.Add(x.ArWorkHours(d), False)
        Next

    End Sub

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        Dim y As New ViewEditWorkHours
        Dim strResponse As String = InputBox$("Enter a new work hour to add to list.", "New work hour", "")
        If strResponse.ToString.Length <= 0 Then
            MsgBox("You cannot have a blank work hour.", MsgBoxStyle.Exclamation, "Error Adding work hour")
            Exit Sub
        ElseIf strResponse.ToString.Length >= 1 Then
            y.InsertHour(strResponse)
            ResetDefault()
            y.PopulateWorkHours()
            Dim d
            For d = 0 To y.ArWorkHours.Count - 1
                Me.chklstWorkHours.Items.Add(y.ArWorkHours(d), False)
            Next
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.ResetDefault()
        Me.Close()
    End Sub
End Class
