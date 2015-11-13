Public Class EmailTemplate


    

    Private Sub EmailTemplate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ResetForm()
    End Sub

    Private Sub txtsubject_GotFocus(sender As Object, e As EventArgs) Handles txtsubject.GotFocus
        Me.lblSubject.Visible = False
    End Sub

    Private Sub txtsubject_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsubject.LostFocus
        Try
            Dim txt As TextBox = sender
            Dim str_txt As String = txt.Text
            If Len(str_txt) <= 0 Then
                Me.lblSubject.Visible = True
            ElseIf Len(str_txt) >= 1 Then
                Me.lblSubject.Visible = False
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtEmailBody_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtfEmailBody.GotFocus
        Me.lblEmailBody.Visible = False
    End Sub

    Private Sub txtEmailBody_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtfEmailBody.LostFocus
        Try
            Dim rtf As RichTextBox = sender
            Dim rtf_txt As String = rtf.Text
            If Len(rtf_txt) <= 0 Then
                Me.lblEmailBody.Visible = True
            ElseIf Len(rtf_txt) >= 1 Then
                Me.lblEmailBody.Visible = False
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        'Me.rtfEmailBody.Text = ""
        'Me.txtsubject.Text = ""
        'Me.txtEmailBody_LostFocus(Nothing, Nothing)
        'Me.txtsubject_LostFocus(Nothing, Nothing)
        ResetForm()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        ResetForm()
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim ctrl As Button = sender
        Dim ctrl_text As String = ctrl.Text

        Select Case ctrl_text
            Case Is = "Save and Close"
                Dim sub_len As Integer = Me.txtsubject.Text.ToString.Length
                If sub_len <= 0 Then
                    MsgBox("You must supply a subject for this template.", MsgBoxStyle.Critical, "No Subject Given")
                    Exit Sub
                End If
                Dim bodyLen As Integer = Me.rtfEmailBody.Text.ToString.Length
                If bodyLen <= 0 Then
                    MsgBox("You must have a body for this email template.", MsgBoxStyle.Critical, "No Message Body")
                    Exit Sub
                End If
                Dim temp_name As String = InputBox("Please enter a name for this template.", "Enter Template Name", "<New Template>")
                Dim tempLen As Integer = temp_name.ToString.Length
                If tempLen <= 0 Then
                    MsgBox("You must supply a template name in order to save it.", MsgBoxStyle.Critical, "No Template Name")
                    Exit Sub
                ElseIf tempLen >= 1 Then
                    '' check for duplicates
                    Dim y As New emlTemplateLogic
                    Dim z As New emlTemplateLogic.TemplateInfo
                    z.ID = 0
                    z.Subject = Me.txtsubject.Text
                    z.Body = Me.rtfEmailBody.Text
                    z.TemplateName = temp_name
                    Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                    z.Department = y.GetEmployeeDepartment(name(0), name(1), False)
                    Dim exists As Boolean = y.CheckDuplicateTemplateExists(False, z)
                    If exists = False Then
                        '' insert it
                        y.InsertNewTemplate(z, False)
                        y = Nothing
                        ResetForm()
                        Confirming.GetTemplates()
                        Me.Close()
                    ElseIf exists = True Then
                        MsgBox("There is already a templated named: " & z.TemplateName & " stored." & vbCrLf & "Please choose another name and try again.", MsgBoxStyle.Information, "Duplicate Exists")
                        Exit Sub
                    End If
                End If
            Case Is = "Update"
                Dim sub_len As Integer = Me.txtsubject.Text.ToString.Length
                If sub_len <= 0 Then
                    MsgBox("You must supply a subject for this template.", MsgBoxStyle.Critical, "No Subject Given")
                    Exit Sub
                End If
                Dim bodyLen As Integer = Me.rtfEmailBody.Text.ToString.Length
                If bodyLen <= 0 Then
                    MsgBox("You must have a body for this email template.", MsgBoxStyle.Critical, "No Message Body")
                    Exit Sub
                End If
                Dim temp_name As String = Me.cboTemplateName.Text
                If Len(temp_name) <= 0 Then
                    MsgBox("Can't update a template without a template name.", MsgBoxStyle.Critical, "Error Updating Template")
                    Exit Sub
                ElseIf Len(temp_name) >= 1 Then
                    Dim zz As New emlTemplateLogic.TemplateInfo
                    zz.ID = "0"
                    zz.Subject = Me.txtsubject.Text
                    zz.Body = Me.rtfEmailBody.Text
                    zz.TemplateName = temp_name
                    Dim xy As New emlTemplateLogic
                    Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                    zz.Department = xy.GetEmployeeDepartment(name(0), name(1), False)
                    xy.UpdateWholeTemplate(False, zz)
                    ResetForm()
                    Confirming.GetTemplates()
                    Me.Close()
                End If
                Exit Select
            Case Else
                Exit Select
        End Select
    End Sub
    Private Sub ResetForm()
        Me.txtsubject.Text = ""
        Me.rtfEmailBody.Text = ""
        Me.cboTemplateName.Text = ""
        Me.btnSave.Text = "Save and Close"
        Me.lblEmailBody.Visible = True
        Me.lblSubject.Visible = True
        Me.lblTemplateName.Visible = True
        Dim y As New emlTemplateLogic
        Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim fname As String = name(0)
        Dim lname As String = name(1)
        Dim b As New List(Of emlTemplateLogic.TemplateInfo)
        b = y.GetTemplatesByDepartment(False, y.GetEmployeeDepartment(fname, lname, False))
        Dim a As emlTemplateLogic.TemplateInfo
        Me.cboTemplateName.Items.Clear()
        For Each a In b
            Me.cboTemplateName.Items.Add(a.TemplateName)
        Next

    End Sub

    Private Sub cboTemplateName_GotFocus(sender As Object, e As EventArgs) Handles cboTemplateName.GotFocus
        Me.lblTemplateName.Visible = False
    End Sub

    Private Sub cboTemplateName_LostFocus(sender As Object, e As EventArgs) Handles cboTemplateName.LostFocus
        Dim cbo As ComboBox = sender
        Dim txt As String = cbo.Text
        If Len(txt) <= 0 Then
            Me.lblTemplateName.Visible = True
            Me.btnSave.Text = "Save and Close"
        ElseIf Len(txt) >= 1 Then
            Me.lblTemplateName.Visible = False
        End If
    End Sub

    Private Sub cboTemplateName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTemplateName.SelectedIndexChanged
       Dim cbo As ComboBox = sender
        Dim sel_str As String = cbo.Text
        If sel_str.ToString.Length >= 1 Then
            Dim z As New emlTemplateLogic
            Dim zz As New emlTemplateLogic.TemplateInfo
            Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            zz = z.GetSingleTemplate(sel_str, False, z.GetEmployeeDepartment(name(0), name(1), False))
            Me.txtsubject.Text = zz.Subject
            Me.rtfEmailBody.Text = zz.Body
            Me.btnSave.Text = "Update"
        ElseIf sel_str.ToString.Length <= 0 Then
            Me.btnSave.Text = "Save and Close"
            Exit Sub
        End If
    End Sub

    Private Sub txtsubject_TextChanged(sender As Object, e As EventArgs) Handles txtsubject.TextChanged
        Dim txt_sub As TextBox = sender
        Dim str_txt As String = txt_sub.Text
        If Len(str_txt) <= 0 Then
            Me.lblSubject.Visible = True
        ElseIf Len(str_txt) >= 1 Then
            Me.lblSubject.Visible = False
        End If
    End Sub

    Private Sub rtfEmailBody_TextChanged(sender As Object, e As EventArgs) Handles rtfEmailBody.TextChanged
        Dim rtf As RichTextBox = sender
        Dim rtf_txt As String = rtf.Text
        If Len(rtf_txt) <= 0 Then
            Me.lblEmailBody.Visible = True
        ElseIf Len(rtf_txt) >= 1 Then
            Me.lblEmailBody.Visible = False
        End If
    End Sub

    Private Sub InsertValueWhere(ByVal Tag As String)
        If Me.txtsubject.Focused = True Then
            Me.txtsubject.Text = (Me.txtsubject.Text & Tag)
            Me.txtsubject.SelectionStart = Len(Me.txtsubject.Text.ToString)
        ElseIf Me.rtfEmailBody.Focused = True Then
            Me.rtfEmailBody.Text = (Me.rtfEmailBody.Text & Tag)
            Me.rtfEmailBody.SelectionStart = Len(Me.rtfEmailBody.Text.ToString)
        ElseIf Me.txtsubject.Focused = False And Me.rtfEmailBody.Focused = False Then
            Exit Sub
        End If
    End Sub

    Private Sub btnCustomer_Click(sender As Object, e As EventArgs) Handles btnCustomer.Click
        InsertValueWhere("<<CustomerName(s)>>")
    End Sub

    Private Sub btnAddress_Click(sender As Object, e As EventArgs) Handles btnAddress.Click
        InsertValueWhere("<<Address>>")
    End Sub

    Private Sub btnCityState_Click(sender As Object, e As EventArgs) Handles btnCityState.Click
        InsertValueWhere("<<CityState>>")
    End Sub

    Private Sub btnZipCode_Click(sender As Object, e As EventArgs) Handles btnZipCode.Click
        InsertValueWhere("<<ZipCode>>")
    End Sub

    Private Sub btnMainPhone_Click(sender As Object, e As EventArgs) Handles btnMainPhone.Click
        InsertValueWhere("<<MainPhone>>")
    End Sub

    Private Sub btnAltPhone1_Click(sender As Object, e As EventArgs) Handles btnAltPhone1.Click
        InsertValueWhere("<<AltPhone1>>")
    End Sub

    Private Sub btnAltPhone2_Click(sender As Object, e As EventArgs) Handles btnAltPhone2.Click
        InsertValueWhere("<<AltPhone2>>")
    End Sub

    Private Sub btnProducts_Click(sender As Object, e As EventArgs) Handles btnProducts.Click
        InsertValueWhere("<<Product(s)>>")
    End Sub

    Private Sub btnSpokeWith_Click(sender As Object, e As EventArgs) Handles btnSpokeWith.Click
        InsertValueWhere("<<SpokeWith>>")
    End Sub

    Private Sub btnApptDate_Click(sender As Object, e As EventArgs) Handles btnApptDate.Click
        InsertValueWhere("<<ApptDate>>")
    End Sub

    Private Sub btnApptTime_Click(sender As Object, e As EventArgs) Handles btnApptTime.Click
        InsertValueWhere("<<ApptTime>>")
    End Sub

    Private Sub btnLastMarketer_Click(sender As Object, e As EventArgs) Handles btnLastMarketer.Click
        InsertValueWhere("<<LastMarketer>>")
    End Sub

    Private Sub btnConfirmer_Click(sender As Object, e As EventArgs) Handles btnConfirmer.Click
        InsertValueWhere("<<Confirmer>>")
    End Sub

    Private Sub btnSalesRep_Click(sender As Object, e As EventArgs) Handles btnSalesRep.Click
        InsertValueWhere("<<SalesRep>>")
    End Sub

    
    
    Private Sub tsbtnCompanyName_Click(sender As Object, e As EventArgs) Handles tsbtnCompanyName.Click
        InsertValueWhere("<<CompanyName>>")
    End Sub

    Private Sub tsbtnCompanyAddress_Click(sender As Object, e As EventArgs) Handles tsbtnCompanyAddress.Click
        InsertValueWhere("<<CompanyAddress>>")
    End Sub

    Private Sub tsbtnCompanyPhone_Click(sender As Object, e As EventArgs) Handles tsbtnCompanyPhone.Click
        InsertValueWhere("<<CompanyPhone>>")
    End Sub

    Private Sub tsbtnCompanyFax_Click(sender As Object, e As EventArgs) Handles tsbtnCompanyFax.Click
        InsertValueWhere("<<CompanyFax>>")
    End Sub

    Private Sub tsbtnCompanyWebsite_Click(sender As Object, e As EventArgs) Handles tsbtnCompanyWebsite.Click
        InsertValueWhere("<<CompanyWebsite>>")
    End Sub

    Private Sub ToolStripLabel4_Click(sender As Object, e As EventArgs) Handles ToolStripLabel4.Click
        InsertValueWhere("<<CompanyAddressMulti>>")
    End Sub
End Class
