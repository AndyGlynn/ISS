Public Class PopulateAF

    Public Sub New(ByVal frm As Form)
        Dim f = frm
        Dim a As New jbPicsAttachedFiles
        Dim xx As New imgLSTS(STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID) '' will have to change to static vars 
        f.imgLst16 = xx.ImgList16
        f.ImgLst32 = xx.ImgList32
        f.ImgLst64 = xx.ImgList64
        f.ImgLst128 = xx.ImgList128
        f.ImgLst256 = xx.ImgList256
        f.lvAttachedFiles.Items.Clear()
        Dim b As Integer
        Dim g = a
        f.lvAttachedFiles.SmallImageList = f.ImgLst32
        Dim cnt As Integer = g.ArFileNameOnly.Count
        For b = 0 To cnt - 1
            Dim xxx As New ListViewItem
            xxx.Text = g.ArFileNameOnly(b).ToString

            f.lvAttachedFiles.Items.Add(xxx)
            '' Me.ListView1.Items.Add(g.ArFileNameOnly(b).ToString)
        Next
        f.lvAttachedFiles.View = View.SmallIcon
        f.lvAttachedFiles.SmallImageList = f.ImgLst32
        Me.AssignImageKeys(f)
        a.Populate_Sub_Directories(f)
    End Sub
    Private Sub AssignImageKeys(ByVal frm As Form)
        Dim f = frm
        Dim bb
        For bb = 0 To f.lvAttachedFiles.Items.Count - 1
            Dim gg As ListViewItem
            gg = f.lvAttachedFiles.Items.Item(bb)
            Dim strName = gg.Text
            Dim FileNameAndExt = Split(strName, ".")
            gg.ImageKey = (FileNameAndExt(1))

        Next
        f.lvAttachedFiles.Refresh()
    End Sub
End Class
