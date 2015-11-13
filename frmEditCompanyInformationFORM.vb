Public Class frmEditCompanyInformation
    Public cmpy_info As CompanyInformation

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim str
        Dim stNum As String = ""
        Dim stAddress As String = ""
        Dim Addy_Extended As String = "" ' IE: Road, Street, Drive, Avenue
        Try
            str = Split(Me.txtStreetAddress.Text, " ", 3)
            stNum = Trim(str(0).ToString)
            stAddress = Trim(str(1).ToString) & " " & Trim(str(2).ToString)
            stAddress = stNum & " " & Trim(str(1).ToString) & " " & Trim(str(2).ToString)
        Catch ex As Exception
            MsgBox("There is a problem with your address format. Please try again. IE: '1234 There st.", MsgBoxStyle.Critical, "Error With Address")
            Exit Sub
        End Try


        cmpy_info.StAddress = stAddress
        cmpy_info.Address_Line_2 = Me.txtAddressLine2.Text
        cmpy_info.State = Me.txtState.Text
        cmpy_info.Zip = Me.txtZip.Text
        cmpy_info.City = Me.txtCity.Text

        cmpy_info.Company_WebSite = Me.txtWebsite.Text
        cmpy_info.Company_Name = Me.txtCompanyName.Text
        cmpy_info.ContactPhoneNumber = Me.txtContactPhoneNumber.Text
        cmpy_info.ContactFaxNumber = Me.txtFaxNumber.Text
        cmpy_info.Logo_Directory = Me.txtLogoDirectory.Text


        cmpy_info.UpdateCompanyInformation()
        Me.ResetDefault()
        cmpy_info.GetInformation()
        '' comment this out after
        '' inserted into production.
        '' 
        'Me.txtStreetAddress.Text = cmpy_info.StAddress
        'Me.txtAddressLine2.Text = cmpy_info.Address_Line_2
        'Me.txtCity.Text = cmpy_info.City
        'Me.txtState.Text = cmpy_info.State
        'Me.txtZip.Text = cmpy_info.Zip
        'Me.txtWebsite.Text = cmpy_info.Company_WebSite
        'Me.txtCompanyName.Text = cmpy_info.Company_Name
        'Me.txtContactPhoneNumber.Text = cmpy_info.ContactPhoneNumber
        'Me.txtFaxNumber.Text = cmpy_info.ContactFaxNumber
        'Me.txtLogoDirectory.Text = cmpy_info.Logo_Directory
        Me.ResetDefault()
        Me.Close()

    End Sub

    Private Sub frmEditCompanyInformation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        cmpy_info = New CompanyInformation
        cmpy_info.GetInformation()

        Me.txtStreetAddress.Text = cmpy_info.StAddress
        Me.txtAddressLine2.Text = cmpy_info.Address_Line_2
        Me.txtCity.Text = cmpy_info.City
        Me.txtState.Text = cmpy_info.State
        Me.txtZip.Text = cmpy_info.Zip
        Me.txtWebsite.Text = cmpy_info.Company_WebSite
        Me.txtCompanyName.Text = cmpy_info.Company_Name
        Me.txtContactPhoneNumber.Text = cmpy_info.ContactPhoneNumber
        Me.txtFaxNumber.Text = cmpy_info.ContactFaxNumber
        Me.txtLogoDirectory.Text = cmpy_info.Logo_Directory

    End Sub
    Private Sub ResetDefault()

        Me.txtStreetAddress.Text = ""
        Me.txtAddressLine2.Text = ""
        Me.txtCity.Text = ""
        Me.txtState.Text = ""
        Me.txtZip.Text = ""
        Me.txtWebsite.Text = ""
        Me.txtCompanyName.Text = ""
        Me.txtContactPhoneNumber.Text = ""
        Me.txtFaxNumber.Text = ""
        Me.txtLogoDirectory.Text = ""

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.ResetDefault()
        Me.Close()
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Me.OpenFileDialog1.ShowDialog()
    End Sub
End Class
