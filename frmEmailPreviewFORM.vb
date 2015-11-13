Public Class frmEmailPreview

    Public LeadToShow As convertLeadToStruct.EnterLead_Record
    Public TemplateName As String = ""
    Public Department As String = ""


    Private Sub frmEmailPreview_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ResetForm()
        Me.cboTemplates.Items.Clear()

    End Sub

    Private Sub frmEmailPreview_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.lblLeadNum.Text = LeadToShow.RecID
        Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim z As New emlTemplateLogic
        Dim depart_ As String = z.GetEmployeeDepartment(name(0), name(1), False)
        Me.Department = depart_
        Dim template_scrubbed As String = z.TestTemplateScrub(LeadToShow.RecID, False, TemplateName, depart_)
        Dim templt_info As emlTemplateLogic.TemplateInfo = z.GetSingleTemplate(TemplateName, False, depart_)
        Dim g As List(Of emlTemplateLogic.TemplateInfo)
        g = z.GetTemplatesByDepartment(False, depart_)
        Dim a As emlTemplateLogic.TemplateInfo
        For Each a In g
            Me.cboTemplates.Items.Add(a.TemplateName)
        Next
        Me.rtfPreview.Text = template_scrubbed
        Me.lblAppliedTemplate.Text = TemplateName.ToString
        Me.lblSubject.Text = z.SubjectScrub(LeadToShow.RecID, False, TemplateName, depart_)
    End Sub

    Private Sub ResetForm()
        LeadToShow = Nothing
        TemplateName = ""
        Me.lblLeadNum.Text = ""
        Me.rtfPreview.Text = ""
        Me.Department = ""
        Me.lblSubject.Text = ""
        Me.lblAppliedTemplate.Text = ""
        Me.cboTemplates.Items.Clear()
        Dim y As New emlTemplateLogic
        Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim depart_ As String = y.GetEmployeeDepartment(name(0), name(1), False)
        Dim g As List(Of emlTemplateLogic.TemplateInfo)
        g = y.GetTemplatesByDepartment(False, depart_)
        Dim a As emlTemplateLogic.TemplateInfo
        For Each a In g
            Me.cboTemplates.Items.Add(a.TemplateName)
        Next
        Me.Name = "Email Preview"
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboTemplates_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTemplates.SelectedIndexChanged
        Dim cbo As ComboBox = sender
        Dim cbo_txt As String = cbo.Text
        Dim z As New emlTemplateLogic
        Dim template_scrubbed As String = z.TestTemplateScrub(LeadToShow.RecID, False, cbo_txt, Me.Department)
        Dim templt_info As emlTemplateLogic.TemplateInfo = z.GetSingleTemplate(cbo_txt, False, Me.Department)
        Me.rtfPreview.Text = template_scrubbed
        Me.lblAppliedTemplate.Text = cbo_txt
        Me.lblSubject.Text = z.SubjectScrub(LeadToShow.RecID, False, cbo_txt, Me.Department)
    End Sub

    Private Sub btnSendEmail_Click(sender As Object, e As EventArgs) Handles btnSendEmail.Click
        Dim y As New EmailIssuedLeads
        y.Send_BLAST_MAIL("aaron.clay79@gmail.com", "aaron.clay79@gmail.com", Me.rtfPreview.Text, Me.lblSubject.Text)
        MsgBox("Email Sent.", MsgBoxStyle.Information, "SENT EMAIL")
        Me.ResetForm()
        Me.Close()
    End Sub
End Class
