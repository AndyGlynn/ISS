Public Class EditCustomerInfo

    Public d As EDIT_CUSTOMER_INFORMATION
    Public ID As String = STATIC_VARIABLES.CurrentID
    Public WhichForm As Form = STATIC_VARIABLES.ActiveChild
    Public Cls As Boolean = False
    Public MP_Verified As Boolean
    Public IsManage As Boolean = False
    Public OAddy As String = ""
    Public OCity As String = ""
    Public OState As String = ""
    Public OZip As String = ""
    Dim chk As Boolean = True
    Dim eCnt As Integer = 0


    Private Sub EditCustomerInfo_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If ID <> "" Then
            Dim c As New Locked_for_Editing
            c.Remove_Lock(ID)
        End If

    End Sub






 

  

  



  
    Private Sub EditCustomerInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      
        If ID = "" Then


            MsgBox("You must select a Record to Edit!", MsgBoxStyle.Exclamation, "No Record Selected")
            Me.Close()
            Exit Sub
        End If
        If WhichForm.Name = "Confirming" Then
            Me.Label5.Text = "Appt. Info is Read Only, To Move this Appt. for this customer Use Move Appt. Button Provided"
        Else
            Me.Label5.Text = "Appt. Info is Read Only, To Set New Appt. for this customer Use Set Appt. Button Provided"
        End If

        ';Me.ResetForm()
        Dim c As New Locked_for_Editing
        c.Lock()
        If Me.Cls = True Then
            Me.Close()
            Exit Sub
        End If
        STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
        If ID <> "" Then
            Me.Text = Me.Text & " for Record ID: " & ID
        End If
        Me.MdiParent = Main
        d = New EDIT_CUSTOMER_INFORMATION
        d.FeedProperties(ID)

        Me.isManager()


        If Me.MP_Verified = False Then
            Me.pctVerified.Visible = True
            Me.btnMap.Enabled = True
        End If
 




    End Sub

    Private Sub ResetForm()
        Me.Text = "Edit Customer Information"
        Me.txtApptDate.Text = ""
        Me.txtContact2.Text = ""
        Me.txtContact1.Text = ""
        Me.txtHousePhone.Text = ""
        Me.cboalt1type.Text = ""
        Me.cboAlt2Type.Text = ""
        Me.txtaltphone2.Text = ""
        Me.txtaltphone1.Text = ""
        Me.cboProduct1.Text = ""
        Me.cboProduct2.Text = ""
        Me.cboProduct3.Text = ""
        Me.rtbSpecialInstructions.Text = ""
        Me.txtApptTime.Text = ""
        Me.txtColor.Text = ""
        Me.txtQty.Text = ""
        Me.txtYrsOwned.Text = ""
        Me.txtYrBuilt.Text = ""
        Me.txtHomeValue.Text = ""
        Me.txtApptDay.Text = ""
        Me.lnkEmail.Text = " "
        Me.ID = ""
        Me.WhichForm = Nothing

    End Sub





 

    Private Sub txtAlt2Type_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim str = Mid(Me.cboAlt2Type.Text, 1, 1)
            Dim str2 = Mid(Me.cboAlt2Type.Text, 2, Me.cboAlt2Type.Text.ToString.Length)

            str = str.ToString.ToUpper
            Dim New_Word As String = str & str2
            Me.cboAlt2Type.Text = New_Word
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAlt1Type_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim str = Mid(Me.cboalt1type.Text, 1, 1)
            Dim str2 = Mid(Me.cboalt1type.Text, 2, Me.cboalt1type.Text.ToString.Length)

            str = str.ToString.ToUpper
            Dim New_Word As String = str & str2
            Me.cboalt1type.Text = New_Word
        Catch ex As Exception

        End Try
    End Sub




    Private Sub EditCustomerInfo_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged

        Me.WindowState = FormWindowState.Normal
    End Sub
    Private Sub cboC1WorkHours_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboC1WorkHours.SelectedValueChanged
        If sender.SelectedItem = "<Add New>" Then
            Dim y As String = ""
            y = sender.SelectedItem
            If y.ToString.Length < 2 Then
                Exit Sub
            End If
            Select Case y
                Case Is = "<Add New>"

                    Dim b As New EDIT_CUSTOMER_INFORMATION
                    Dim pr As String = ""


                    pr = b.CapitalizeText(InputBox$("Enter ""Work Hours"" name", "New Work Hours"))
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
    Private Sub cboProduct1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct1.SelectedValueChanged
        If Me.cboProduct1.SelectedItem = "<Add New>" Then
            Dim y As String = ""
            y = Me.cboProduct1.SelectedItem
            If y.ToString.Length < 2 Then
                Exit Sub
            End If
            Select Case y
                Case Is = "<Add New>"

                    Dim b As New EDIT_CUSTOMER_INFORMATION
                    Dim pr As String = ""
                    Dim prA As String = ""

                    pr = b.CapitalizeText(InputBox$("Enter new product name", "New Product Name"))
                    If pr.ToString.Length < 2 Then
                        Me.cboProduct1.Text = ""
                        Exit Sub
                    End If
                    prA = b.CapitalizeText(InputBox$("Enter product acronym.(2 Letters)", "Product Acronym"))
                    If prA.ToString.Length < 2 Then
                        Me.cboProduct1.Text = ""
                        Exit Sub
                    End If
                    b.AddNewProduct(pr, prA, "CBO1")
                    Exit Select
                Case Is = ""
                    Exit Select
                Case Is = "___________________________________________"
                    Me.cboProduct1.Text = ""
                    Exit Select
            End Select
        End If
    End Sub


    Private Sub cboProduct2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct2.SelectedValueChanged
      
        If Me.cboProduct2.SelectedItem = "<Add New>" Then
            Dim y As String = ""
            y = Me.cboProduct2.SelectedItem
            If y.ToString.Length < 2 Then
                Exit Sub
            End If
            Select Case y
                Case Is = "<Add New>"

                    Dim b As New EDIT_CUSTOMER_INFORMATION
                    Dim pr As String = ""
                    Dim prA As String = ""

                    pr = b.CapitalizeText(InputBox$("Enter new product name", "New Product Name"))
                    If pr.ToString.Length < 2 Then
                        Me.cboProduct1.Text = ""
                        Exit Sub
                    End If
                    prA = b.CapitalizeText(InputBox$("Enter product acronym.(2 Letters)", "Product Acronym"))
                    If prA.ToString.Length < 2 Then
                        Me.cboProduct1.Text = ""
                        Exit Sub
                    End If
                    b.AddNewProduct(pr, prA, "CBO2")
                    Exit Select
                Case Is = ""
                    Exit Select
                Case Is = "___________________________________________"
                    Me.cboProduct2.Text = ""
                    Exit Select
            End Select
        End If
    End Sub

    Private Sub cboProduct3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProduct3.SelectedValueChanged
        If Me.cboProduct3.SelectedItem = "<Add New>" Then
            Dim y As String = ""
            y = Me.cboProduct3.SelectedItem
            If y.ToString.Length < 2 Then
                Exit Sub
            End If
            Select Case y
                Case Is = "<Add New>"

                    Dim b As New EDIT_CUSTOMER_INFORMATION
                    Dim pr As String = ""
                    Dim prA As String = ""

                    pr = b.CapitalizeText(InputBox$("Enter new product name", "New Product Name"))
                    If pr.ToString.Length < 2 Then
                        Me.cboProduct1.Text = ""
                        Exit Sub
                    End If
                    prA = b.CapitalizeText(InputBox$("Enter product acronym.(2 Letters)", "Product Acronym"))
                    If prA.ToString.Length < 2 Then
                        Me.cboProduct1.Text = ""
                        Exit Sub
                    End If
                    b.AddNewProduct(pr, prA, "CBO3")
                    Exit Select
                Case Is = ""
                    Exit Select
                Case Is = "___________________________________________"
                    Me.cboProduct3.Text = ""
                    Exit Select
            End Select
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        'Select Case STATIC_VARIABLES.ActiveChild.Name
        '    Case "Confirming"
        '        Exit Select
        '    Case "WCaller"
        '        Exit Select
        Me.Chk_Errors(Me.CurrentTab)
        If Me.eCnt >= 1 Then
            Exit Sub
        End If
        d.New_Properties()
        Dim x As New which_form(WhichForm, ID)
        Me.ResetForm()
        Me.Close()





    End Sub

  

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
    Private CurrentTab As Integer = 0
  

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        ' Me.ToolTip1.RemoveAll()
        Select Case Me.TabControl1.SelectedIndex
       
            Case 5
                If Me.IsManage = False Then
                    'Me.ToolTip1.Show("", Me.TabPage5)
                    Me.ToolTip1.Show("You don't have permission to access the Lead Source Tab", Me.TabControl1, 5000)
                    Me.TabControl1.SelectedIndex = Me.CurrentTab
                    Exit Select
                End If

     
        End Select


    End Sub

    Private Sub Chk_Errors(ByVal tab)
        Dim c As New EDIT_CUSTOMER_INFORMATION
        ep.Clear()
        eCnt = 0
        Select Case Me.CurrentTab
            Case 0
                If Me.txtContact2.Text = "" Then
                    Me.cboC2WorkHours.SelectedItem = ""
                End If
                If Me.txtaltphone1.Text = "" Then
                    Me.cboalt1type.SelectedItem = ""
                End If
                If Me.txtaltphone2.Text = "" Then
                    Me.cboAlt2Type.Text = ""
                End If

                If Me.txtContact1.Text = "" Then
                    ep.SetError(Me.txtContact1, "Required Field")
                    eCnt += 1
                End If
                Dim cnt As Integer = 0
                Dim i As Char
                For Each i In Me.txtHousePhone.Text
                    Select Case i
                        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
                            cnt += 1
                    End Select
                Next

                If cnt <> 10 Then
                    ep.SetError(Me.txtHousePhone, "Required Field")
                    eCnt += 1
                End If
                If Me.txtContact2.Text <> "" And Me.cboC2WorkHours.Text = "" Then
                    ep.SetError(Me.cboC2WorkHours, "Required Field")
                    eCnt += 1
                End If
                If Me.cboC1WorkHours.Text = "" Then
                    ep.SetError(Me.cboC1WorkHours, "Required Field")
                    eCnt += 1
                End If
                If Me.txtContact1.Text.Contains(" ") = False Then
                    ep.SetError(Me.txtContact1, "Last Name Required")
                    eCnt += 1
                Else
                    If c.Get_Last(Me.txtContact1.Text) = "" Then
                        ep.SetError(Me.txtContact1, "Last Name Required")
                        eCnt += 1
                    End If
                End If
                If eCnt >= 1 Then
                    Me.TabControl1.SelectedIndex = 0
                End If
            Case 1
                If Me.txtAddress.Text = "" Then
                    ep.SetError(Me.txtAddress, "Required Field")
                    eCnt += 1
                End If
                If Me.txtCity.Text = "" Then
                    ep.SetError(Me.txtCity, "Required Field")
                    eCnt += 1
                End If
                If Me.txtState.Text = "" Then
                    ep.SetError(Me.txtState, "Required Field")
                    eCnt += 1
                End If
                If eCnt >= 1 Then
                    Me.TabControl1.SelectedIndex = 1
                End If
            Case 3
                If Me.cboProduct1.SelectedItem = Nothing Then
                    ep.SetError(Me.cboProduct1, "Required Field")
                    eCnt += 1
                End If
            Case 5
                If Me.cboPriLead.SelectedItem = Nothing Then
                    ep.SetError(Me.cboPriLead, "Required Field")
                    eCnt += 1
                End If
                If eCnt >= 1 Then
                    Me.TabControl1.SelectedIndex = 5
                End If
        End Select

    End Sub
    Private Sub isManager()
        Dim cnt As Integer = 0
        If STATIC_VARIABLES.Administration = True Then
            cnt += 1
        End If
        If STATIC_VARIABLES.Install = True Then
            cnt += 1
        End If
        If STATIC_VARIABLES.Finance = True Then
            cnt += 1
        End If
        If STATIC_VARIABLES.MarketingManager = True Then
            cnt += 1
        End If
        If STATIC_VARIABLES.SalesManager = True Then
            cnt += 1
        End If
        If cnt >= 1 Then
            Me.IsManage = True
        End If
    End Sub
    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        Me.Chk_Errors(Me.CurrentTab)
        Me.CurrentTab = Me.TabControl1.SelectedIndex
    End Sub

    Private Sub txtContact2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContact2.LostFocus
        Dim c As New EDIT_CUSTOMER_INFORMATION

        If Me.txtContact2.Text.Contains(" ") = False Then
            Me.txtContact2.Text = Me.txtContact2.Text & " " & c.Get_Last(Me.txtContact1.Text)
        End If
    End Sub

    Private Sub cboPriLead_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPriLead.SelectedValueChanged
        Dim c As String = ""
        c = Me.cboPriLead.Text
        If c.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case c
            Case Is = "<Add New>"
                Dim f As New EDIT_CUSTOMER_INFORMATION

                Dim pri As String = ""
                pri = f.CapitalizeText(InputBox$("Enter new Primary Lead Source.", "New Primary Lead Source"))
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
                Dim d As New EDIT_CUSTOMER_INFORMATION
                d.GetSLS(c)
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
                Dim b As New EDIT_CUSTOMER_INFORMATION
                Dim sec As String = ""
                Dim cap As New ENTER_LEAD.Captilalize
                sec = cap.CapitalizeText((InputBox$("Enter new Secondary Lead Source.", "New Secondary Lead Source")))
                If sec = "" Then
                    Me.cboSecLead.Text = ""
                    Exit Sub
                End If
                b.InsertSLS(y, sec)

                b.GetSLS(y)
                Me.cboSecLead.SelectedItem = sec
                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "_____________________________________________"
                Me.cboSecLead.Text = ""
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

                Dim c As New EDIT_CUSTOMER_INFORMATION
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

                    d.GetMarketerLeadSources(fname, lname)

                Catch ex As Exception

                End Try
                Exit Select
        End Select
    End Sub
    Dim click As Integer
    Private Sub btnSetAppt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetAppt2.Click
        If WhichForm.Name = "Confirming" Then
            RescheduleAppt.frm = WhichForm.Name
            RescheduleAppt.ID = ID
            RescheduleAppt.ShowDialog()
            d.Get_New_Appt(ID)
            Me.txtApptDate.Text = d.AppointmentDate
            Me.txtApptTime.Text = d.AppointmentTime
            Me.txtApptDay.Text = d.AppointmentDay
            Exit Sub
        End If
        click += 1
        SetAppt.OrigApptDate = Me.txtApptDate.Text
        SetAppt.OrigApptTime = Me.txtApptTime.Text
        SetAppt.Contact1 = d.Get_First(Me.txtContact1.Text)
        SetAppt.Contact2 = d.Get_First(Me.txtContact2.Text)
        SetAppt.frm = WhichForm
        'If click = 1 Then
        '    SetAppt.ShowDialog()
        'End If

        SetAppt.ShowDialog() ''come back weird fuckin bug, form opens and closes first time only then works fine 
        d.Get_New_Appt(ID)


    End Sub

    Private Sub btnMap_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMap.Click
        Dim c As New Edit_Verify_Address(Me.txtAddress.Text, Me.txtCity.Text, Me.txtState.Text, Me.txtZip.Text, 1)

    End Sub

    Private Sub txtAddress_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddress.LostFocus
        chk = Me.Chk_Cng
        If chk = False And Me.btnMap.Enabled = False Then
            Me.btnMap.Enabled = True
            Me.pctVerified.Visible = True
            d.Update_MPVerified_False(ID)
        End If

    End Sub
    Function Chk_Cng() As Boolean
        Dim chk1 As Boolean = True
        If Me.txtAddress.Text <> Me.OAddy Then
            chk1 = False
        End If
        If Me.txtCity.Text <> Me.OCity Then
            chk1 = False
        End If
        If Me.txtState.Text <> Me.OState Then
            chk1 = False
        End If
        If Me.txtZip.Text <> Me.OZip Then
            chk1 = False
        End If
        Return chk1
    End Function

    Private Sub txtCity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCity.LostFocus
        chk = Me.Chk_Cng
        If chk = False And Me.btnMap.Enabled = False Then
            Me.btnMap.Enabled = True
            Me.pctVerified.Visible = True
            d.Update_MPVerified_False(ID)
        End If
    End Sub

    Private Sub txtState_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtState.LostFocus
        chk = Me.Chk_Cng
        If chk = False And Me.btnMap.Enabled = False Then
            Me.btnMap.Enabled = True
            Me.pctVerified.Visible = True
            d.Update_MPVerified_False(ID)
        End If
    End Sub

    Private Sub txtZip_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtZip.LostFocus
        chk = Me.Chk_Cng
        If chk = False And Me.btnMap.Enabled = False Then
            Me.btnMap.Enabled = True
            Me.pctVerified.Visible = True
            d.Update_MPVerified_False(ID)
        End If
    End Sub
    

    Private Sub cboC2WorkHours_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboC2WorkHours.SelectedValueChanged
        If sender.SelectedItem = "<Add New>" Then
            Dim y As String = ""
            y = sender.SelectedItem
            If y.ToString.Length < 2 Then
                Exit Sub
            End If
            Select Case y
                Case Is = "<Add New>"

                    Dim b As New EDIT_CUSTOMER_INFORMATION
                    Dim pr As String = ""


                    pr = b.CapitalizeText(InputBox$("Enter ""Work Hours"" name", "New Work Hours"))
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
