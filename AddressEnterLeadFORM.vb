Public Class AddressEnterLead
    Public frm As Form
    Public IsThisAddressVerified As Boolean '' will double up to send back to the insertlead class
    Public ForceMappedToTrue As Boolean '' will be for the Do Not Update flag
    Public StopProcessing As Boolean '' variable if "go back" was hit to end the rest of the processing
    Public SelectedAddress As ListViewItem
    Public Update_StAddress As String = ""
    Public Update_City As String = ""
    Public Update_State As String = ""
    Public Update_Zip As String = ""
    Public USE_UPDATE_ADDRESS_INSTEAD As Boolean = False
    Private Sub AddressEnterLead_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Label1.Text.Contains(STATIC_VARIABLES.ProgramName) Then
            Label1.Text = Replace(Label1.Text, STATIC_VARIABLES.ProgramName, "")
        End If
        ' Open as Show Dialog
        'Uncomment the Code below when you get this into your project
        Me.Label1.Text = STATIC_VARIABLES.ProgramName & Me.Label1.Text
        EnterLead.ExSub = False

    End Sub

    Private Sub lvAddresses_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvAddresses.MouseDown
        SelectedAddress = Me.lvAddresses.GetItemAt(e.X, e.Y)
        
    End Sub

  
 
    Private Sub lvAddresses_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvAddresses.SelectedIndexChanged
        If Me.lvAddresses.SelectedItems.Count > 0 Then
            
            
            Me.btnUpdate.Enabled = True
        Else
            Me.btnUpdate.Enabled = False
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        If Me.frm.Name = "Enterlead" Then
            Me.StopProcessing = True
            Me.lvAddresses.Items.Clear()

            EnterLead.ExSub = True

            EnterLead.txtStAddy.SelectAll()
            EnterLead.txtStAddy.Focus()
            Me.Close()
        ElseIf Me.frm.Name = "EditCustomerInfo" Then

        End If
    
        
    End Sub

    Private Sub btnDoNotUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDoNotUpdate.Click
      
        'EnterLead.MapPoint_Verified = False
        'Dim x As New VerifyAddress("", "", "", "", 4)
        If Me.frm.Name = "Enterlead" Then
            EnterLead.MapPoint_Verified = False
            EnterLead.ExSub = False
            Me.Close()
            Me.Dispose()
            Me.StopProcessing = False
            Me.ForceMappedToTrue = False
        ElseIf Me.frm.Name = "EditCustomerInfo" Then

        End If



        ''Go back to enterlead @ finish executing Save & New or Save & Close
        '' Write lead to table with @mapped = to False 
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        '' Split Address and replace address paramaters with new and write @mapped = true 
        ''finish executing Save & New or Save & Close
        Dim ADDRESS() = Split(SelectedAddress.Text, ",", 3) '' first get the splits at the commas
        Dim STA As String = ADDRESS(0)
        Dim CITY As String = ADDRESS(1)
        Dim STATEZIP As String = ADDRESS(2)
        Dim ST As String = ""
        Dim ZP As String = ""
        Dim STATEANDZIP = Split(Trim(STATEZIP), " ", 2)
        ST = Trim(STATEANDZIP(0))
        ZP = Trim(STATEANDZIP(1))
        Me.Update_City = CITY
        Me.Update_State = ST
        Me.Update_Zip = ZP
        Me.Update_StAddress = STA
        'EnterLead.txtCity.Text = Me.Update_City
        'EnterLead.txtState.Text = Me.Update_State
        'EnterLead.txtZip.Text = Me.Update_Zip
        'EnterLead.txtStAddy.Text = Me.Update_StAddress
        'Me.USE_UPDATE_ADDRESS_INSTEAD = True
        'EnterLead.MapPoint_Verified = True

        If Me.frm.Name = "Enterlead" Then
            Dim x As New VerifyAddress(Me.Update_StAddress, Me.Update_City, Me.Update_State, Me.Update_Zip, 1)
        ElseIf Me.frm.Name = "EditCustomerInfo" Then
            Dim x As New Edit_Verify_Address(Me.Update_StAddress, Me.Update_City, Me.Update_State, Me.Update_Zip, 1)
        End If

        Me.Close()
    End Sub

End Class
