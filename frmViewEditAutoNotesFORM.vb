Public Class frmViewEditAutoNotes

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.ResetDefault()
        Me.Close()
    End Sub
    Private Sub ResetDefault()
        Me.chklstAutoNotes.Items.Clear()
        'Me.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ResetDefault()
        Dim x As New ViewEditAutoNotes
        x.PopuluateList()

        Dim d
        For d = 0 To x.ArAutoNotes.Count - 1
            Me.chklstAutoNotes.Items.Add(x.ArAutoNotes(d), False)
        Next

    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim y As New ViewEditAutoNotes
            Dim t As Object
            Dim str As String = ""
            For Each t In Me.chklstAutoNotes.CheckedItems
                str = t.ToString
                If str.ToString.Length <= 0 Then
                    '' do nothing
                ElseIf str.ToString.Length >= 1 Then
                    y.DeleteNote(str)
                End If
            Next
            ResetDefault()
            y.PopuluateList()
            Dim d
            For d = 0 To y.ArAutoNotes.Count - 1
                Me.chklstAutoNotes.Items.Add(y.ArAutoNotes(d), False)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim y As New ViewEditAutoNotes
        Dim strResponse As String = InputBox$("Enter a new auto note to add to list.", "New Auto Note", "")
        If strResponse.ToString.Length <= 0 Then
            MsgBox("You cannot have a blank auto note.", MsgBoxStyle.Exclamation, "Error Adding Auto Note")
            Exit Sub
        ElseIf strResponse.ToString.Length >= 1 Then
            y.InsertAutoNote(strResponse)
            ResetDefault()
            y.PopuluateList()
            Dim d
            For d = 0 To y.ArAutoNotes.Count - 1
                Me.chklstAutoNotes.Items.Add(y.ArAutoNotes(d), False)
            Next
        End If
    End Sub
End Class
