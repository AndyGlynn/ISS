

Public Class ImportPictures
    Private ProdAcro As String = ""
    Private Sub ImportPictures_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim z As New IMPORT_PICTURES_LOGIC
        z.GetProducts()
        Me.txtLeadNum.Text = STATIC_VARIABLES.CurrentID.ToString
        'Dim b
        'For b = 1 To z.ProdCnt
        '    Me.cboProductSel.Items.Add(z.prods(b).ToString)
        'Next
        'Me.txtLeadNum.Text = RecordLogic.CurrentID.ToString
    End Sub



    Private Sub cboProductSel_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProductSel.SelectedValueChanged
        Dim x As String = Me.cboProductSel.Text
        Dim z As New IMPORT_PICTURES_LOGIC
        z.PullACRO(x.ToString)
        ProdAcro = z.ProductAcronym
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Dim chkrdos As Integer
        chkrdos = 0
        If Me.txtLeadNum.Text = "" Then
            Me.ErrorProvider1.SetError(Me.txtLeadNum, "You must supply a Record ID.")
            Exit Sub
        End If
        If Me.cboProductSel.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboProductSel, "You must supply a product.")
            Exit Sub
        End If
        Dim x
        LeadNum = ""
        If Me.rdoBefore.Checked = True Then
            x = "B"
            chkrdos = 1
            ProgressVal = x
            LeadNum = Me.txtLeadNum.Text

        End If
        If Me.rdoMiddle.Checked = True Then
            x = "M"
            chkrdos = 1
            ProgressVal = x
            LeadNum = Me.txtLeadNum.Text
        End If
        If Me.rdoAfter.Checked = True Then
            x = "A"
            chkrdos = 1
            ProgressVal = x
            LeadNum = Me.txtLeadNum.Text
        End If
        Select Case chkrdos
            Case Is = 1
                'ValidateID.Validate()
                'Select Case ValidateID.ThereFlag
                '    Case True
                OpenPics.Multiselect = True

                OpenPics.Title = "Select Picture(s) to Import"
                OpenPics.Filter = "JPEG|*.jpg|BMP|*.bmp|TIFF|*.tif|PNG|*.png|WMF|*.wmf|GIF|*.gif|All Files|*.*"
                OpenPics.FileName = "[Select a File]"
                OpenPics.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop '"C:\documents and settings\All users\Desktop"
                OpenPics.ShowDialog()

                Import.FileName = Me.txtLeadNum.Text & "-" & ProgressVal & "-" & ProdAcro.ToString
                Import.StripGarbage()

                Me.txtLeadNum.Text = ""
                Me.cboProductSel.Text = ""

                'Case False
                'MsgBox("Our Database cannot find Lead # " & Me.txtLeadNum.Text & ". Please check your entry and try again.", MsgBoxStyle.Exclamation, "Error Validating Lead Number")
                'Me.txtLeadNum.Text = ""
                'Me.txtLeadNum.Select()
                'Exit Sub
                Exit Select
            Case Is < 1
                Me.ErrorProvider1.SetError(Me.GroupBox1, "You must select job picture status.")
                Exit Select
        End Select
        ProdAcro = ""
        Me.Close()
    End Sub



    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()

    End Sub
End Class
