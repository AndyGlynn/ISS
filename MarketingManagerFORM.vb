Public Class MarketingManager

    Private Sub MarketingManager_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = Main
        Dim y As New ScheduledActions
        y.SetUp(Me)

    End Sub

    Private Sub btnMarkTaskAsDone_Click(sender As Object, e As EventArgs) Handles btnMarkTaskAsDone.Click
        Dim toggle As String = ""
        Dim btn As ToolStripMenuItem = sender
        Dim strLen As Integer = 0
        toggle = btn.Text.ToString
        strLen = toggle.ToString.Length

        Select Case strLen
            Case Is <> 24 '' All Else / Mark as Done
                For Each x As Control In Me.pnlScheduledTasks.Controls
                    If TypeOf x Is Panel Then
                        Dim pnlTarget As Panel = x
                        If pnlTarget.BorderStyle = BorderStyle.FixedSingle Then
                            Dim ID As String = Replace(pnlTarget.Name, "pnl", "")
                            Dim z As New ScheduledActions
                            z.Completed(ID)
                            Me.pnlScheduledTasks.Controls.Clear()
                            z.SetUp(Me)
                        End If
                    End If
                Next
                Exit Select
            Case Is = 24 '' Undo "Mark Task as Done"
                For Each x As Control In Me.pnlScheduledTasks.Controls
                    If TypeOf x Is Panel Then
                        Dim pnlTarget As Panel = x
                        If pnlTarget.BorderStyle = BorderStyle.FixedSingle Then
                            Dim ID As String = Replace(pnlTarget.Name, "pnl", "")
                            Dim z As New ScheduledActions
                            z.UndoCompleted(ID)
                            Me.pnlScheduledTasks.Controls.Clear()
                            z.SetUp(Me)
                        End If
                    End If
                Next
                Exit Select
        End Select
    End Sub

    Private Sub btnEditScheduledTask_Click(sender As Object, e As EventArgs) Handles btnEditScheduledTask.Click
        ScheduleAction.edit = True
        Dim c As Integer = Me.pnlScheduledTasks.Controls.Count
        Dim i As Integer
        For i = 1 To c
            Dim all As Panel = Me.pnlScheduledTasks.Controls(i - 1)
            If all.BorderStyle = BorderStyle.FixedSingle Then
                ScheduleAction.EditId = all.Name.ToString.Substring(3)
            End If
        Next
        If ScheduleAction.EditId = "" Then
            MsgBox("You must select a Task.", MsgBoxStyle.Exclamation, "Please Select a Task")
            Exit Sub
        End If
        ScheduleAction.ShowDialog()
        Dim x As New ScheduledActions
        x.SetUp(Me)
    End Sub

    Private Sub btnRemoveThisCompletedTask_Click(sender As Object, e As EventArgs) Handles btnRemoveThisCompletedTask.Click
        For Each x As Control In Me.pnlScheduledTasks.Controls
            If TypeOf x Is Panel Then
                Dim pnlTarget As Panel = x
                If pnlTarget.BorderStyle = BorderStyle.FixedSingle Then
                    Dim ID As String = Replace(pnlTarget.Name, "pnl", "")
                    Dim z As New ScheduledActions
                    z.HideTask(ID)
                    Me.pnlScheduledTasks.Controls.Clear()
                    z.SetUp(Me)
                End If
            End If
        Next
    End Sub

    Private Sub btnRemoveAllScheduledTask_Click(sender As Object, e As EventArgs) Handles btnRemoveAllScheduledTask.Click
        Dim x As New ScheduledActions
        x.HideAll(Me)
        Me.pnlScheduledTasks.Controls.Clear()
        x.SetUp(Me)
    End Sub

    Private Sub tsCreateList_Click(sender As Object, e As EventArgs) Handles tsCreateList.Click
        frmCreateList.Show()
        frmCreateList.MdiParent = Main

    End Sub
End Class
