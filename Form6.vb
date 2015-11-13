Public Class Form6
    Public pbMemMax
    Public pbRowMax
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        If Me.btnImport.Text = "Close" Then
            Me.Close()
        Else
            frmImportBulkData.txtDirectory.Text = Me.txtDirectory.Text
        End If

    End Sub


    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtDirectory_TextChanged(sender As Object, e As EventArgs) Handles txtDirectory.TextChanged
        frmImportBulkData.Visible = False
        frmImportBulkData.Show()
        frmImportBulkData.Visible = False
        frmImportBulkData.txtDirectory.Text = Me.txtDirectory.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        Me.FolderBrowserDialog1.ShowDialog()
        Dim path As String = Me.FolderBrowserDialog1.SelectedPath
        Me.txtDirectory.Text = path
    End Sub

   
    Private Sub lblTable_Click(sender As Object, e As EventArgs) Handles lblTable.Click

    End Sub

    Private Sub lblTable_VisibleChanged(sender As Object, e As EventArgs) Handles lblTable.VisibleChanged
        ' MsgBox(lblTable.Visible.ToString)
        'If Me.lblTable.Visible = False Then
        '    Me.lblTotalRecordNum.Visible = False
        '    Me.lblTblCurCnt.Visible = False
        '    Me.lblCurrentRecordNum.Visible = False
        'Else
        '    Me.lblTotalRecordNum.Visible = True
        '    Me.lblTblCurCnt.Visible = True
        '    Me.lblCurrentRecordNum.Visible = True
        'End If
    End Sub
End Class
