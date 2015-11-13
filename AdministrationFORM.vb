Public Class Administration



    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      
    End Sub

    Private Sub TreeView1_BeforeSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs)

    End Sub

    Private Sub TreeView1_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs)

    End Sub

    Private Sub TreeView1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)


    End Sub

    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs)
        ' MsgBox(Me.TreeView1.SelectedNode.Text)
        'If Me.TreeView1.Nodes(0).Checked = True Then
        '    Me.TreeView1.Nodes(0).Nodes(0).Checked = True
        '    Me.TreeView1.Nodes(0).Nodes(0).Nodes(0).Checked = True
        '    Me.TreeView1.Nodes(0).Nodes(0).Nodes(1).Checked = True
        '    Me.TreeView1.Nodes(0).Nodes(0).Nodes(2).Checked = True
        'ElseIf Me.TreeView1.Nodes(0).Checked = False Then
        '    Me.TreeView1.Nodes(0).Nodes(0).Checked = False
        '    Me.TreeView1.Nodes(0).Nodes(0).Nodes(0).Checked = False
        '    Me.TreeView1.Nodes(0).Nodes(0).Nodes(1).Checked = False
        '    Me.TreeView1.Nodes(0).Nodes(0).Nodes(2).Checked = False
        'End If
    End Sub

    Private Sub TreeView1_QueryAccessibilityHelp(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryAccessibilityHelpEventArgs)

    End Sub

    Private Sub tsbtnUsers_Click(sender As Object, e As EventArgs) Handles tsbtnUsers.Click
        SetUpUser.ShowDialog()

    End Sub

    Private Sub Administration_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
