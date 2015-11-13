'Imports MapPoint
'Imports System.Diagnostics.Process


Public Class EnterLead

    Public MapPoint_Verified As Boolean '' mappoint variable to make sure flow of verifyaddress isn't interupted
    Private epEnterLeadSwitch As Boolean '' error provider switch to return true/false depending if anything was caught // someone didn't fill in a required feild. (Spellcheck)
    Public ExSub As Boolean = False


    Private Sub EnterLead_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Me.BackgroundWorker1_DoWork(Nothing, Nothing)
        If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
            STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
        End If
        Me.epEnterLead.Clear() '' clear off old values passed to error provider
        Dim c As New ENTER_LEAD
        c.Loadup()
        Me.dtpApptInfo.Value = Today.AddDays(1)
        Me.txtApptTime.Value = CType("01/01/1900 " & Now.Hour.ToString & ":00", Date)
    End Sub

    Private Sub cboPriLead_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPriLead.SelectedValueChanged
    
        Dim c As String = ""
        c = Me.cboPriLead.Text
        If c.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case c
            Case Is = "<Add New>"
                Dim f As New ENTER_LEAD.InsertPLS
                Dim pri As String = ""
                pri = InputBox$("Enter new Primary Lead Source.", "New Primary Lead Source")
                If pri.ToString.Length < 2 Then
                    Me.cboPriLead.Text = ""
                    Me.cboSecLead.Items.Clear()
                    Exit Sub
                End If
                f.InsertNewPLS(pri)
                'Dim rq As New ENTER_LEAD.PopulatePrimaryLeadSource
                'rq.GetPrimaryLeadSource()
                'Me.cboPriLead.SelectedItem = pri
                Exit Select
            Case Is = ""
                Me.cboSecLead.Items.Clear()
                Exit Select
            Case Is = "_____________________________________________"
                Me.cboPriLead.Text = ""
                Me.cboSecLead.Items.Clear()
                Exit Select
            Case Else
                Dim d As New ENTER_LEAD.PopulateSecondaryLeadSource
                d.GetSLS(c)
                Exit Select
        End Select
        
    End Sub

   

    Private Sub cboMarketer_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMarketer.SelectedValueChanged
        Dim x As String = ""
        x = Me.cboMarketer.Text
        If x.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case x
            Case Is = "<Add New>"
                '' add new marketer roll

                Dim c As New ENTER_LEAD.InsertMarketer
                c.InsertMarketer(Me)
            Case Is = ""
                Me.cboMarketer.Text = ""
                Dim c As String = Me.cboMarketer.Text
                If c = "" Then
                    Me.cboPriLead.Enabled = True
                    Me.cboPriLead.Text = ""
                    'Dim d As New ENTER_LEAD.PopulatePrimaryLeadSource
                    'd.GetPrimaryLeadSource()
                    Me.cboSecLead.Enabled = True
                    Me.cboSecLead.Items.Clear()
                    Me.cboSecLead.Text = ""
                End If
                Exit Select
            Case Is = "_____________________________________________"
                Me.cboMarketer.Text = ""
                Exit Select
            Case Else
                Try
                    Dim fname As String = ""
                    Dim lname As String = ""
                    Dim name
                    name = Split(x, " ", 2)
                    fname = name(0)
                    lname = name(1)
                    Dim d As New ENTER_LEAD.PopulatePLSandSLSbyMarketer
                    d.GetMarketerLeadSources(fname, lname)
                 
                Catch ex As Exception

                End Try
                Exit Select
        End Select
    End Sub

    Private Sub txtApptTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As String = Me.txtApptTime.Text
        If b.ToString.Length < 2 Then
            Exit Sub
        End If
        Dim tf As New TimeFormat
        tf.CheckTimeFormat(b)
        Me.txtApptTime.Text = tf.RetTime
    End Sub

    Private Sub dtpApptInfo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptInfo.LostFocus
        'Dim c As New ENTER_LEAD
        'c.GetDayOfWeek()
    End Sub

    Private Sub dtpApptInfo_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptInfo.ValueChanged
        '' On load, this causes an infinite loop 
        '' value changed event is calling a NEW enter_lead class
        '' inside of a new enter_lead class
        '' 3/26/08 AC
        Dim d1 = Me.dtpApptInfo.Value.DayOfWeek
        Select Case d1
            Case Is = 0
                Me.txtApptday.Text = "Sunday"
                Exit Select
            Case Is = 1
                Me.txtApptday.Text = "Monday"
                Exit Select
            Case Is = 2
                Me.txtApptday.Text = "Tuesday"
                Exit Select
            Case Is = 3
                Me.txtApptday.Text = "Wednesday"
                Exit Select
            Case Is = 4
                Me.txtApptday.Text = "Thursday"
                Exit Select
            Case Is = 5
                Me.txtApptday.Text = "Friday"
                Exit Select
            Case Is = 6
                Me.txtApptday.Text = "Saturday"
                Exit Select
        End Select

    End Sub

    Private Sub cboSecLead_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSecLead.SelectedValueChanged
        Dim x As String = ""
        x = Me.cboSecLead.Text
        Dim y As String = ""
        y = Me.cboPriLead.Text
        If y.ToString.Length < 2 Then
            Exit Sub
        End If
        If x.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case x
            Case Is = "<Add New>"
                Dim b As New ENTER_LEAD.InsertSLS
                Dim sec As String = ""
                Dim cap As New ENTER_LEAD.Captilalize
                sec = cap.CapitalizeText((InputBox$("Enter new Secondary Lead Source.", "New Secondary Lead Source")))
                If sec = "" Then
                    Me.cboSecLead.Text = ""
                    Exit Sub
                End If
                b.InsertSLS(y, sec)
                Dim d As New ENTER_LEAD.PopulateSecondaryLeadSource
                d.GetSLS(y)
                Me.cboSecLead.SelectedItem = sec
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_____________________________________________"
                Me.cboSecLead.Text = ""
                Exit Select
        End Select
    End Sub

    Private Sub cboProduct1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct1.SelectedValueChanged
        Dim y As String = ""
        y = Me.cboProduct1.SelectedItem
        If y.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case y
            Case Is = "<Add New>"
                Dim b As New ENTER_LEAD.InsertProduct
                Dim pr As String = ""
                Dim prA As String = ""
                Dim cap As New ENTER_LEAD.Captilalize
                pr = cap.CapitalizeText(InputBox$("Enter new product name", "New Product Name"))
                If pr.ToString.Length < 2 Then
                    Me.cboProduct1.Text = ""
                    Exit Sub
                End If
                prA = cap.CapitalizeText(InputBox$("Enter product acronym.(2 Letters)", "Product Acronym"))
                If prA.ToString.Length < 2 Then
                    Me.cboProduct1.Text = ""
                    Exit Sub
                End If
                b.AddNewProduct(pr, prA, "CBO1")
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_________________________"
                Me.cboProduct1.Text = ""
                Exit Select
        End Select
    End Sub

    Private Sub cboProduct2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct2.SelectedIndexChanged
        Dim str As String = Me.cboProduct2.Text
        Select Case str
            Case Is = "<Add New>"
                Exit Select
            Case Is = " "
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_________________________"
                Me.cboProduct2.Text = ""
                Exit Select
            Case Else
                Dim x As New ENTER_LEAD
                x.GetAcronym(Me.cboProduct2.Text, 2)
                Exit Select
        End Select
    End Sub

    Private Sub cboProduct2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct2.SelectedValueChanged
        Dim y As String = ""
        y = Me.cboProduct2.SelectedItem
        If y.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case y
            Case Is = "<Add New>"
                Dim b As New ENTER_LEAD.InsertProduct
                Dim pr As String = ""
                Dim prA As String = ""
                Dim cap As New ENTER_LEAD.Captilalize
                pr = cap.CapitalizeText(InputBox$("Enter new product name", "New Product Name"))
                If pr.ToString.Length < 2 Then
                    Me.cboProduct2.Text = ""
                    Exit Sub
                End If
                prA = cap.CapitalizeText(InputBox$("Enter product acronym.(2 Letters)", "Product Acronym"))
                If prA.ToString.Length < 2 Then
                    Me.cboProduct2.Text = ""
                    Exit Sub
                End If
                b.AddNewProduct(pr, prA, "CBO2")
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_________________________"
                Me.cboProduct2.Text = ""
                Exit Select
        End Select
        If Me.cboProduct2.Text <> "" Then
            Dim x As New ENTER_LEAD
            x.GetAcronym(Me.cboProduct2.Text, 2)
        End If
       
    End Sub

    Private Sub cboProduct3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct3.SelectedIndexChanged
        Dim str As String = Me.cboProduct3.Text
        Select Case str
            Case Is = "<Add New>"
                Exit Select
            Case Is = " "
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_________________________"
                Me.cboProduct3.Text = ""
                Exit Select
            Case Else
                Dim x As New ENTER_LEAD
                x.GetAcronym(Me.cboProduct3.Text, 3)
                Exit Select
        End Select
    End Sub

    Private Sub cboProduct3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct3.SelectedValueChanged
        Dim y As String = ""
        y = Me.cboProduct3.SelectedItem
        If y.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case y
            Case Is = "<Add New>"
                Dim b As New ENTER_LEAD.InsertProduct
                Dim pr As String = ""
                Dim prA As String = ""
                Dim cap As New ENTER_LEAD.Captilalize
                pr = cap.CapitalizeText(InputBox$("Enter new product name", "New Product Name"))
                If pr.ToString.Length < 2 Then
                    Me.cboProduct3.Text = ""
                    Exit Sub
                End If
                prA = cap.CapitalizeText(InputBox$("Enter product acronym. (2 Letters)", "Product Acronym"))
                If prA.ToString.Length < 2 Then
                    Me.cboProduct3.Text = ""
                    Exit Sub
                End If
                b.AddNewProduct(pr, prA, "CBO3")
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_________________________"
                Me.cboProduct3.Text = ""
                Exit Select
        End Select
        If Me.cboProduct3.Text <> "" Then
            Dim x As New ENTER_LEAD
            x.GetAcronym(Me.cboProduct3.Text, 3)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.cboMarketer.SelectedItem = "Will Smith"
        Me.cboProduct1.SelectedItem = "Roof"
        Me.cboProduct2.SelectedItem = "Windows"
        Me.cboProduct3.SelectedItem = "Siding"
        Me.cboC1Work.SelectedItem = "1st Shift"
        Me.cboC2Work.SelectedItem = "2nd Shift"
        Me.cboAlt1Type.Text = "Work"
        Me.cboAltPhone2Type.Text = "Cell"
        Me.txtC1FName.Text = "Test"
        Me.txtC1LName.Text = "Insert"
        Me.txtC2FName.Text = "Test"
        Me.txtC2LName.Text = "Insert2"
        Me.txtHousePhone.Text = "4194724000"
        Me.txtAltPhone1.Text = "4194666984"
        Me.txtAltPhone2.Text = "4194723373"
        Me.txtProdColor.Text = "Brown"
        Me.txtProdQTY.Text = "3"
        Me.txtYearsOwned.Text = "25"
        Me.txtAgeOfHome.Text = "25"
        Me.txtHomeVal.Text = "125"
        Me.txtApptTime.Text = "1:00 PM"
        Me.rtfSpecialIns.Text = "Test Insert"
        Me.txtStAddy.Text = "1512 Michigan Ave"
        'Me.txtStNum.Text = "1512"
        Me.txtCity.Text = "Maumee"
        Me.txtZip.Text = "43537"
        Me.txtState.Text = "OH"
        Me.txtEmail.Text = "SomeGuy@there.com"

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub txtC1FName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtC1FName.LostFocus
        Dim y As New ENTER_LEAD.Captilalize
        Me.txtC1FName.Text = y.CapitalizeText(Me.txtC1FName.Text)
        Me.cboSpokeWith.Items.Clear()
        Me.cboSpokeWith.Items.Add(Me.txtC1FName.Text)
    End Sub

    Private Sub txtC2FName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtC2FName.LostFocus
        '' method fails in a scenario of only 1 char entered for this field.
        '' IE: 'A' or a first initial // skips over logic because IF statement doesn't compensate for chars < 2.
        '' Is anyone ever going to just use a first initial for a contact 'First' name?
        ''
        Dim y As New ENTER_LEAD.Captilalize
        Dim z As String = ""
        z = Me.txtC2FName.Text
        If z.ToString.Length < 2 Then
            Exit Sub
        End If
        If z.ToString.Length > 2 Then
            Me.txtC2FName.Text = y.CapitalizeText(Me.txtC2FName.Text)
            Me.cboSpokeWith.Items.Clear()
            Me.cboSpokeWith.Items.Add(Me.txtC1FName.Text)
            Me.cboSpokeWith.Items.Add(Me.txtC2FName.Text)
            Me.cboSpokeWith.Items.Add(Me.txtC1FName.Text & " and " & Me.txtC2FName.Text)
            'Me.cboSpokeWith.Items.Add("Both")
        End If
    End Sub

    Private Sub txtC1LName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtC1LName.LostFocus
        'Dim g As New ENTER_LEAD.Captilalize
        'Me.txtC1FName.Text = g.CapitalizeText(Me.txtC1FName.Text)
        Dim text As String = Me.txtC1LName.Text
        Me.txtC1LName.Text = StrConv(text, VbStrConv.ProperCase)
    End Sub

    Private Sub txtC2LName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtC2LName.LostFocus
        '' first find out if there is a name in txtc1lname
        '' if there is text present there, and this feild is skipped over
        '' add the last name assuming that they are married or what have you
        '' then capitalize it.

        Dim len As Integer = 0
        len = Me.txtC1LName.Text.Length
        Select Case len
            Case Is <= 0
                Exit Sub
            Case Is >= 1

                Dim len2 As Integer = 0
                len2 = Me.txtC2FName.Text.Length
                '' if there is no first name in c2fname
                '' exit out of here
                If len2 <= 0 Then
                    Exit Select
                End If

                '' if there is a first name, check to see if 
                '' there is a last name present in c2lname
                If len2 >= 1 Then
                    Dim len3 As Integer = 0
                    len3 = Me.txtC2LName.Text.Length
                    If len3 <= 0 Then
                        Me.txtC2LName.Text = StrConv(Me.txtC1LName.Text, VbStrConv.ProperCase)
                        Me.cboC2Work.Text = ""
                    End If
                    If len3 >= 1 Then
                        Me.txtC2LName.Text = StrConv(Me.txtC2LName.Text, VbStrConv.ProperCase)
                        Me.cboC2Work.Text = ""
                    End If
                End If
        End Select

        '' now check if there is even a name present here.
        '' if no name, changed contact2 work hours to "N/A"

        Dim lenC2FName As Integer = 0
        Dim lenC2LName As Integer = 0
        lenC2FName = Me.txtC2FName.Text.Length
        lenC2LName = Me.txtC2LName.Text.Length
        If (lenC2FName + lenC2LName) = 0 Then
            Me.cboC2Work.Text = "N/A"
        End If

        '' now check to see if both first name of contact1 
        '' and first name of contact2 present. if they are,
        '' string them together in cboSpokeWith

        'Dim lenC1 As Integer = 0
        'Dim lenC2 As Integer = 0
        'lenC1 = Me.txtC1FName.Text.Length
        'lenC2 = Me.txtC2FName.Text.Length
        'If lenC1 And lenC2 >= 1 Then
        '    Dim C1 As String = Me.txtC1FName.Text
        '    Dim c2 As String = Me.txtC2FName.Text
        '    Me.cboSpokeWith.Items.Add(C1 & " and " & c2)
        'End If

    End Sub

    Private Sub btnSaveNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveNew.Click
        '' 3.30.08 the error provider needs to catch the form before it even goes through the rest of its logic.
        '' NEED a method here for the errorprovider to bounce in and out of.
        '' must be replicated on Save as well.....
        ExSub = False
        Me.epEnterLead.Clear()
        Me.CheckRequiredFields()
        If Me.epEnterLeadSwitch = True Then
            Exit Sub
        End If
        If Me.epEnterLeadSwitch = False Then
            Dim c As New VerifyAddress(Me.txtStAddy.Text, Me.txtCity.Text, Me.txtState.Text, Me.txtZip.Text, 1)
            'Dim verified As Boolean = c.MapPointVerified
            'Me.MapPoint_PASSBACK = verified
            '' if the user has chose to go back,
            '' here is the logic to flip the switch back so that it continues last time
            '' and resets the cursor to the StAddy feild on EnterLead Form
            ' ''
            'If AddressEnterLead.StopProcessing = True Then
            '    AddressEnterLead.StopProcessing = False
            '    DuplicateLead.CloseMethod = "SAVEANDNEW"
            '    Me.Focus()
            '    Me.txtStAddy.Focus()
            '    Exit Sub
            'End If
            '' likewise, if the user has decided to push their address through
            '' logic needs to go here to reflect that change
            ''
            'If AddressEnterLead.StopProcessing = False Then
            '    If AddressEnterLead.ForceMappedToTrue = False Then
            '        If AddressEnterLead.USE_UPDATE_ADDRESS_INSTEAD = False Then
            '            DuplicateLead.CloseMethod = "SAVEANDNEW"
            '            Dim STA As String = ""
            '            Dim CTY As String = ""
            '            Dim ST As String = ""
            '            Dim ZP As String = ""

            '            STA = AddressEnterLead.Update_StAddress
            '            CTY = AddressEnterLead.Update_City
            '            ST = AddressEnterLead.Update_State
            '            ZP = AddressEnterLead.Update_Zip

            '            'Dim b As New ENTER_LEAD.CheckDuplicateLead(STA, CTY, ST, ZP, "SAVEANDNEW", AddressEnterLead.ForceMappedToTrue) // CLAY
            '            Dim b As New ENTER_LEAD.PopulateDuplicates
            '            b.SetUp(Me.txtC1FName.Text, Me.txtC1LName.Text, Me.txtC2FName.Text, Me.txtC2LName.Text, Me.txtHousePhone.Text, Me.txtAltPhone1.Text, Me.txtAltPhone2.Text, STA, CTY, ST, ZP, "SAVEANDNEW", Me.MapPoint_Verified)

            '        End If
            'End If
            'End If
            '' need a method here to compensate for a scenario where the address is correct and it just needs to run through.
            ''
            '' Verify Address still needs to kick back what it found to this part
            '' IE: One address found, address is correct, now just save and close the form.
            ''
            'If AddressEnterLead.StopProcessing = False Then
            '    If AddressEnterLead.ForceMappedToTrue = True Then
            '        If AddressEnterLead.USE_UPDATE_ADDRESS_INSTEAD = True Then
            '            DuplicateLead.CloseMethod = "SAVE"
            '            Dim sta As String = ""
            '            Dim cty As String = ""
            '            Dim st As String = ""
            '            Dim zp As String = ""

            '            sta = AddressEnterLead.Update_StAddress
            '            cty = AddressEnterLead.Update_City
            '            st = AddressEnterLead.Update_State
            '            zp = AddressEnterLead.Update_Zip

            '            Dim b As New ENTER_LEAD.PopulateDuplicates
            '            b.SetUp(Me.txtC1FName.Text, Me.txtC1LName.Text, Me.txtC2FName.Text, Me.txtC2LName.Text, Me.txtHousePhone.Text, Me.txtAltPhone1.Text, Me.txtAltPhone2.Text, sta, cty, st, zp, "SAVEANDNEW", )
            '        End If
            '    End If
            'End If

            '' or the third switch going here to just edit the address
            '' with one of the addresses comming out of mappoint and turning the @Mapped flag to true.
            ' ''
            'If AddressEnterLead.StopProcessing = False Then
            '    If AddressEnterLead.USE_UPDATE_ADDRESS_INSTEAD = True Then
            '        DuplicateLead.CloseMethod = "SAVEANDNEW"
            '        Dim STA As String = ""
            '        Dim CTY As String = ""
            '        Dim ST As String = ""
            '        Dim ZP As String = ""

            '        STA = AddressEnterLead.Update_StAddress
            '        CTY = AddressEnterLead.Update_City
            '        ST = AddressEnterLead.Update_State
            '        ZP = AddressEnterLead.Update_Zip

            'Dim b As New ENTER_LEAD.CheckDuplicateLead(STA, CTY, ST, ZP, "SAVEANDNEW", True) // CLAY
            If Me.ExSub = True Then
                Me.txtStAddy.SelectAll()
                Me.txtStAddy.Focus()
                Exit Sub
            End If
            Dim b As New ENTER_LEAD.PopulateDuplicates
            b.SetUp(Me.txtC1FName.Text, Me.txtC1LName.Text, Me.txtC2FName.Text, Me.txtC2LName.Text, Me.txtHousePhone.Text, Me.txtAltPhone1.Text, Me.txtAltPhone2.Text, Me.txtStAddy.Text, Me.txtCity.Text, Me.txtState.Text, Me.txtZip.Text, "SAVEANDNEW", Me.MapPoint_Verified)

            'AddressEnterLead.StopProcessing = False
            'AddressEnterLead.ForceMappedToTrue = False
            'AddressEnterLead.USE_UPDATE_ADDRESS_INSTEAD = False
            'AddressEnterLead.SelectedAddress = Nothing
            'AddressEnterLead.Update_Zip = ""
            'AddressEnterLead.Update_State = ""
            'AddressEnterLead.Update_City = ""
            'AddressEnterLead.Update_StAddress = ""
            'AddressEnterLead.lvAddresses.Items.Clear()
            '    End If
            'End If
            DuplicateLead.CloseMethod = "SAVEANDNEW"
        End If
    End Sub

    Private Sub btnSaveClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveClose.Click
        '' 3.30.08 the error provider needs to catch the form before it even goes through the rest of its logic.
        '' NEED a method here for the errorprovider to bounce in and out of.
        '' must be replicated on Save and New as well.....
        ''
        '' first clear out old values on error provider if there are any present // REDUNDANCY
        ExSub = False
        Me.epEnterLead.Clear()
        Me.CheckRequiredFields()
        If Me.epEnterLeadSwitch = True Then
            Exit Sub
        End If
        If Me.epEnterLeadSwitch = False Then

            '' NOTES On Error Provider Enter Lead---
            '' Two ways to do this, 1) check all fields at once or 2) check field by field.
            '' first method would catch all at once and illuminate the error provider icon,
            '' second method would provide a step down approach.....IE: caught one, run script, caught second one, run script
            '' vs. caught field x,y,z,a,2 - run script. 
            '' Feild by feild approach would let the user know right then and there on a "lost focus" scenario that they are expected to 
            '' provide information. however i think that would be very annoying to a user especially if they didn't fill in 
            '' a number of fields....
            '' Also, need a list of required fields. I know this was changed so I'll have to get an updated list of what
            '' information this is supposed to be checking/validating.
            '' 
            '' EnterLead [Form] error provider is going to encompass:
            '' Primary Lead Source
            '' Contact1First and Last name
            '' Contact2First and Last name
            '' House Phone
            '' StAddress
            '' City, State
            '' Spoke with
            '' Contact1Work
            '' Product 1
            ''
            '' ALSO: per Andy - Validate all feilds at once or exit sub before rest of logic ensues. 4/1/08
            ''
            '' need to revise these steps here to reflect the additional changes that were made in saveandnew
            ''

            Dim c As New VerifyAddress(Me.txtStAddy.Text, Me.txtCity.Text, Me.txtState.Text, Me.txtZip.Text, 1)
            'Dim verified As Boolean = c.MapPointVerified
            'Me.MapPoint_PASSBACK = verified
            '' if the user has chose to go back,
            '' here is the logic to flip the switch back so that it continues last time
            '' and resets the cursor to the StAddy feild on EnterLead Form
            ' ''
            'If AddressEnterLead.StopProcessing = True Then
            '    AddressEnterLead.StopProcessing = False
            '    DuplicateLead.CloseMethod = "SAVE"
            '    Me.Focus()
            '    Me.txtStAddy.Focus()
            '    Exit Sub
            'End If
            ' '' likewise, if the user has decided to push their address through
            ' '' logic needs to go here to reflect that change
            ' ''
            'If AddressEnterLead.StopProcessing = False Then
            '    If AddressEnterLead.ForceMappedToTrue = False Then
            '        If AddressEnterLead.USE_UPDATE_ADDRESS_INSTEAD = False Then
            '            DuplicateLead.CloseMethod = "SAVE"
            '            Dim STA As String = ""
            '            Dim CTY As String = ""
            '            Dim ST As String = ""
            '            Dim ZP As String = ""

            '            STA = AddressEnterLead.Update_StAddress
            '            CTY = AddressEnterLead.Update_City
            '            ST = AddressEnterLead.Update_State
            '            ZP = AddressEnterLead.Update_Zip

            '            'Dim b As New ENTER_LEAD.CheckDuplicateLead(STA, CTY, ST, ZP, "SAVEANDNEW", AddressEnterLead.ForceMappedToTrue) // CLAY
            '            Dim b As New ENTER_LEAD.PopulateDuplicates
            '            b.SetUp(Me.txtC1FName.Text, Me.txtC1LName.Text, Me.txtC2FName.Text, Me.txtC2LName.Text, Me.txtHousePhone.Text, Me.txtAltPhone1.Text, Me.txtAltPhone2.Text, STA, CTY, ST, ZP, "SAVE", Me.MapPoint_Verified)

            '        End If
            '    End If
            'End If

            ' '' or the third switch going here to just edit the address
            ' '' with one of the addresses comming out of mappoint and turning the @Mapped flag to true.
            ' ''
            'If AddressEnterLead.StopProcessing = False Then
            '    If AddressEnterLead.USE_UPDATE_ADDRESS_INSTEAD = True Then
            '        DuplicateLead.CloseMethod = "SAVE"
            '        Dim STA As String = ""
            '        Dim CTY As String = ""
            '        Dim ST As String = ""
            '        Dim ZP As String = ""

            '        STA = AddressEnterLead.Update_StAddress
            '        CTY = AddressEnterLead.Update_City
            '        ST = AddressEnterLead.Update_State
            '        ZP = AddressEnterLead.Update_Zip
            If Me.ExSub = True Then
                Me.txtStAddy.SelectAll()
                Me.txtStAddy.Focus()
                Exit Sub
            End If
            'Dim b As New ENTER_LEAD.CheckDuplicateLead(STA, CTY, ST, ZP, "SAVE", True) '// CLAY
            Dim b As New ENTER_LEAD.PopulateDuplicates
            b.SetUp(Me.txtC1FName.Text, Me.txtC1LName.Text, Me.txtC2FName.Text, Me.txtC2LName.Text, Me.txtHousePhone.Text, Me.txtAltPhone1.Text, Me.txtAltPhone2.Text, Me.txtStAddy.Text, Me.txtCity.Text, Me.txtState.Text, Me.txtZip.Text, "SAVE", Me.MapPoint_Verified)

            '        AddressEnterLead.StopProcessing = False
            '        AddressEnterLead.ForceMappedToTrue = False
            '        AddressEnterLead.USE_UPDATE_ADDRESS_INSTEAD = False
            '        AddressEnterLead.SelectedAddress = Nothing
            '        AddressEnterLead.Update_Zip = ""
            '        AddressEnterLead.Update_State = ""
            '        AddressEnterLead.Update_City = ""
            '        AddressEnterLead.Update_StAddress = ""
            '        AddressEnterLead.lvAddresses.Items.Clear()
            '    End If
            'End If
            DuplicateLead.CloseMethod = "SAVE"
        End If
    End Sub
    Public Sub Reset()
        Me.cboMarketer.Text = ""
        Me.cboPriLead.Text = ""
        Me.cboSecLead.Text = ""


        Me.txtStAddy.Text = ""
        Me.txtCity.Text = ""
        Me.txtState.Text = ""
        Me.txtZip.Text = ""

        Me.txtC1FName.Text = ""
        Me.txtC1LName.Text = ""
        Me.txtC2FName.Text = ""
        Me.txtC2LName.Text = ""

        Me.txtHousePhone.Text = ""
        Me.txtAltPhone1.Text = ""
        Me.txtAltPhone2.Text = ""
        Me.cboAlt1Type.Text = ""
        Me.cboAltPhone2Type.Text = ""

        Me.txtHomeVal.Text = ""
        Me.txtAgeOfHome.Text = ""
        Me.txtYearsOwned.Text = ""

        Me.rtfSpecialIns.Text = ""

        Me.cboSpokeWith.Text = ""

        Me.cboC1Work.Text = ""
        Me.cboC2Work.Text = ""

        Me.txtApptday.Text = ""
        Me.txtApptTime.Text = ""

        Me.cboProduct1.Text = ""
        Me.cboProduct2.Text = ""
        Me.cboProduct3.Text = ""

        Me.txtProdQTY.Text = ""
        Me.txtProdColor.Text = ""

        Me.txtEmail.Text = ""
    End Sub

    Private Sub cboProduct1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProduct1.SelectedIndexChanged
        Dim str As String = Me.cboProduct1.Text
        Select Case str
            Case Is = "<Add New>"
                Exit Select
            Case Is = " "
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_________________________"
                Me.cboProduct1.Text = ""
                Exit Select
            Case Else
                Dim x As New ENTER_LEAD
                x.GetAcronym(Me.cboProduct1.Text, 1)
                Exit Select
        End Select
    End Sub

    Private Function CheckRequiredFields()
        Dim cnt As Integer = 0 '' variable to count chars in required fields
        '' REQUIRED FIELDS
        '' pls
        '' c1 both
        '' c2 both
        '' HousePhone
        '' StAddress
        '' City
        '' State
        '' Spoke w/
        '' C1 Work
        '' Product 1
        Me.epEnterLead.Clear() '' super redundant call to clear error provider. iterated like 5x.

        If Me.cboPriLead.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.cboPriLead, "Required Field")
        End If
        If Me.txtC1FName.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.txtC1FName, "Required Field")
        End If
        If Me.txtC1LName.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.txtC1LName, "Required Field")
        End If
        'If Me.txtC2FName.Text.Length <= 0 Then
        '    cnt += 1
        '    Me.epEnterLead.SetError(Me.txtC2FName, "Required Field")
        Dim cntp As Integer = 0
        Dim i As Char
        For Each i In Me.txtHousePhone.Text
            Select Case i
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
                    cntp += 1
            End Select
        Next

        If cntp <> 10 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.txtHousePhone, "Required Field")
        End If
      
        If Me.txtStAddy.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.txtStAddy, "Required Field")
        End If
        If Me.txtCity.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.txtCity, "Required Field")
        End If
        If Me.txtState.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.txtState, "Required Field")
        End If
        If Me.cboSpokeWith.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.cboSpokeWith, "Required Field")
        End If
        If Me.cboC1Work.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.cboC1Work, "Required Field")
        End If
        If Me.cboProduct1.Text.Length <= 0 Then
            cnt += 1
            Me.epEnterLead.SetError(Me.cboProduct1, "Required Field")
        End If
        '' do i need a count of the vars if i am returning a switch to be processed by another method?
        '' guess so. works with it.
        If cnt >= 1 Then
            Me.epEnterLeadSwitch = True
        ElseIf cnt <= 0 Then
            Me.epEnterLeadSwitch = False
        End If
        Return Me.epEnterLeadSwitch
    End Function

  
    Private Sub EnterLead_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub txtCity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCity.LostFocus
        If Me.txtCity.Text = "" Then
            Exit Sub
        End If
        Dim x As New ENTER_LEAD
        x.AutoFillState(Me.txtCity.Text)
        If Me.txtState.Text <> "" Then
            Me.txtState.TabStop = False
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.txtStAddy.SelectAll()
        Me.txtStAddy.Focus()
    End Sub

   
    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'STATIC_VARIABLES.oApp = New MapPoint.Application
    End Sub

    Private Sub cboC1Work_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboC1Work.SelectedValueChanged
        If sender.SelectedItem = "<Add New>" Then
            Dim y As String = ""
            y = sender.SelectedItem
            If y.ToString.Length < 2 Then
                Exit Sub
            End If
            Select Case y
                Case Is = "<Add New>"
                    Dim b As New ENTER_LEAD.PopulateWorkHours
                    Dim x As New ENTER_LEAD.Captilalize
                    Dim pr As String = ""


                    pr = x.CapitalizeText(InputBox$("Enter ""Work Hours"" name", "New Work Hours"))
                    If pr.ToString.Length < 2 Then
                        sender.Text = ""
                        Exit Sub
                    End If

                    b.InsertWH(pr, sender)
                    Exit Select
                Case Is = ""
                    Exit Select
                Case Is = "___________________________________________"
                    sender.Text = ""
                    Exit Select
            End Select
        End If
    End Sub


    Private Sub cboC2Work_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboC2Work.SelectedValueChanged
        If sender.SelectedItem = "<Add New>" Then
            Dim y As String = ""
            y = sender.SelectedItem
            If y.ToString.Length < 2 Then
                Exit Sub
            End If
            Select Case y
                Case Is = "<Add New>"
                    Dim b As New ENTER_LEAD.PopulateWorkHours
                    Dim x As New ENTER_LEAD.Captilalize
                    Dim pr As String = ""


                    pr = x.CapitalizeText(InputBox$("Enter ""Work Hours"" name", "New Work Hours"))
                    If pr.ToString.Length < 2 Then
                        sender.Text = ""
                        Exit Sub
                    End If

                    b.InsertWH(pr, sender)
                    Exit Select
                Case Is = ""
                    Exit Select
                Case Is = "___________________________________________"
                    sender.Text = ""
                    Exit Select
            End Select
        End If
    End Sub
End Class

