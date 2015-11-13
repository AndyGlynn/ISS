Imports System.Runtime.InteropServices
Public Class Open_Windows
    Public Sub Win_List()
        Dim count As Integer = Main.MdiChildren.Length
        Main.WindowsToolStripMenuItem.DropDownItems.Clear()


        For i As Integer = 0 To count - 1
            'Dim count As Integer = Me.MdiChildren.Length
            'windowMenu.Items.ClearAndDisposeItems()

            'For i As Integer = 0 To count - 1
            '    frmButton = New C1.Win.C1Ribbon.RibbonButton
            '    frmButton.Text = Me.MdiChildren(i).Text
            '    frmButton.Tag = i
            '    If MdiChildren(i) Is ActiveMdiChild Then
            '        frmButton.SmallImage = My.Resources.test
            '    End If
            '    windowMenu.Items.Add(frmButton)
            '    AddHandler frmButton.Click, AddressOf frmButton_Click
            'Next
           



            Dim frmButton = New ToolStripDropDownButton
            frmButton.Text = Main.MdiChildren(i).Text
            frmButton.Tag = i


            ' Main.WindowsToolStripMenuItem.DropDownItems.Add(frmButton)
            If Main.MdiChildren(i) Is Main.ActiveMdiChild Then
                Main.WindowsToolStripMenuItem.DropDownItems.Add(frmButton)
                Dim w As Integer

               
                'Main.WindowsToolStripMenuItem.
            End If
           
        Next
    End Sub



End Class
