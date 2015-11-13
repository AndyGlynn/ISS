Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports System.IO
Imports System.Text
Imports System.Drawing


Public Class EDIT_CUSTOMER_INFORMATION
#Region "Connection String"
    Private cnn As SqlConnection = New SqlConnection(static_variables.cnn)
#End Region
#Region "Variables"
    '' Variable to select wether the over state of the object has changed
    '' boolean
    Public Data_Has_Changed As Boolean
    ''
    '' contact names
    ''
    Private c1_fname As String = ""
    Private c1_lname As String = ""
    Private c2_fname As String = ""
    Private c2_lname As String = ""
    '' phones 
    ''
    Private main_phone As String = ""
    Private alt1_phone As String = ""
    Private alt1_type As String = ""
    Private alt2_phone As String = ""
    Private alt2_type As String = ""
    '' address
    '' 
    Private StAddy As String = ""
    Private Cty As String = ""
    Private _State As String = ""
    Private _Zip As String = ""
    '' work hours
    '' 
    Private c1_work_hours As String = ""
    Private c2_work_hours As String = ""
    '' email
    '' 
    Private email_addy As String = ""
    '' home information 
    '' 
    Private years_owned As String = ""
    Private home_value As String = ""
    Private year_built As String = ""
    '' Appointment Info
    '' 
    Private Appt_date As String = ""
    Private Appt_time As String = ""
    Private Appt_day As String = ""
    '' Product Info
    '' 
    Private p1 As String = ""
    Private p2 As String = ""
    Private p3 As String = ""
    Private p_qty As String = ""
    Private p_color As String = ""
    '' special instructions
    '' 
    Private spec_instr As String = ""

    '' private vars for functions
    ''
    Private corrected_House_value As String = ""
    Private corrected_Appointment_Time As String = ""
    Private corrected_appointment_date As String = ""

    '' Marketing Info 

    Private Marketer As String = ""
    Private PLeadSource As String = ""
    Private SLeadSource As String = ""
    Private MarketManage As String = ""
#End Region
#Region "Properties"
    Public Property CorrectedHouseValue() As String
        Get
            Return corrected_House_value
        End Get
        Set(ByVal value As String)
            corrected_House_value = value
        End Set
    End Property
    Public Property CorrectedAppointmentTime() As String
        Get
            Return corrected_Appointment_Time
        End Get
        Set(ByVal value As String)
            corrected_Appointment_Time = value
        End Set
    End Property
    Public Property CorrectedAppointmentDate()
        Get
            Return corrected_appointment_date
        End Get
        Set(ByVal value)
            corrected_appointment_date = value
        End Set
    End Property
    Public Property Contact1FirstName() As String
        Get
            Return c1_fname
        End Get
        Set(ByVal value As String)
            c1_fname = value
        End Set
    End Property
    Public Property Contact1LastName() As String
        Get
            Return c1_lname
        End Get
        Set(ByVal value As String)
            c1_lname = value
        End Set
    End Property
    Public Property Contact2FirstName() As String
        Get
            Return c2_fname
        End Get
        Set(ByVal value As String)
            c2_fname = value
        End Set
    End Property
    Public Property Contact2LastName() As String
        Get
            Return c2_lname
        End Get
        Set(ByVal value As String)
            c2_lname = value
        End Set
    End Property
    Public Property MainPhone() As String
        Get
            Return main_phone
        End Get
        Set(ByVal value As String)
            main_phone = value
        End Set
    End Property
    Public Property AltPhone1() As String
        Get
            Return alt1_phone
        End Get
        Set(ByVal value As String)
            alt1_phone = value
        End Set
    End Property
    Public Property AltPhone2() As String
        Get
            Return alt2_phone
        End Get
        Set(ByVal value As String)
            alt2_phone = value
        End Set
    End Property
    Public Property Alt1Type() As String
        Get
            Return alt1_type
        End Get
        Set(ByVal value As String)
            alt1_type = value
        End Set
    End Property
    Public Property Alt2Type() As String
        Get
            Return alt2_type
        End Get
        Set(ByVal value As String)
            alt2_type = value
        End Set
    End Property
    Public Property StAddress() As String
        Get
            Return StAddy
        End Get
        Set(ByVal value As String)
            StAddy = value
        End Set
    End Property
    Public Property City() As String
        Get
            Return Cty
        End Get
        Set(ByVal value As String)
            Cty = value
        End Set
    End Property
    Public Property STATE() As String
        Get
            Return _State
        End Get
        Set(ByVal value As String)
            _State = value
        End Set
    End Property
    Public Property Zip() As String
        Get
            Return _Zip
        End Get
        Set(ByVal value As String)
            _Zip = value
        End Set
    End Property
    Public Property Contact1WorkHours() As String
        Get
            Return c1_work_hours
        End Get
        Set(ByVal value As String)
            c1_work_hours = value
        End Set
    End Property
    Public Property Contact2WorkHours() As String
        Get
            Return c2_work_hours
        End Get
        Set(ByVal value As String)
            c2_work_hours = value
        End Set
    End Property
    Public Property EmailAddress() As String
        Get
            Return email_addy
        End Get
        Set(ByVal value As String)
            email_addy = value
        End Set
    End Property
    Public Property YearsOwned() As String
        Get
            Return years_owned
        End Get
        Set(ByVal value As String)
            years_owned = value
        End Set
    End Property
    Public Property HomeValue() As String
        Get
            Return home_value
        End Get
        Set(ByVal value As String)
            home_value = value
        End Set
    End Property
    Public Property YearBuilt() As String
        Get
            Return year_built
        End Get
        Set(ByVal value As String)
            year_built = value
        End Set
    End Property
    Public Property AppointmentDate() As String
        Get
            Return Appt_date
        End Get
        Set(ByVal value As String)
            Appt_date = value
        End Set
    End Property
    Public Property AppointmentTime() As String
        Get
            Return Appt_time
        End Get
        Set(ByVal value As String)
            Appt_time = value
        End Set
    End Property
    Public Property AppointmentDay() As String
        Get
            Return Appt_day
        End Get
        Set(ByVal value As String)
            Appt_day = value
        End Set
    End Property
    Public Property Product1() As String
        Get
            Return p1
        End Get
        Set(ByVal value As String)
            p1 = value
        End Set
    End Property
    Public Property Product2() As String
        Get
            Return p2
        End Get
        Set(ByVal value As String)
            p2 = value
        End Set
    End Property
    Public Property Product3() As String
        Get
            Return p3
        End Get
        Set(ByVal value As String)
            p3 = value
        End Set
    End Property
    Public Property Product_Quantity() As String
        Get
            Return p_qty
        End Get
        Set(ByVal value As String)
            p_qty = value
        End Set
    End Property
    Public Property Product_Color() As String
        Get
            Return p_color
        End Get
        Set(ByVal value As String)
            p_color = value
        End Set
    End Property
    Public Property Special_Instruction() As String
        Get
            Return spec_instr
        End Get
        Set(ByVal value As String)
            spec_instr = value
        End Set
    End Property
    Public Property Marketer_() As String
        Get
            Return Marketer
        End Get
        Set(ByVal value As String)
            Marketer = value
        End Set
    End Property
    Public Property PriLS() As String
        Get
            Return PLeadSource
        End Get
        Set(ByVal value As String)
            PLeadSource = value
        End Set
    End Property
    Public Property SecLS() As String
        Get
            Return SLeadSource
        End Get
        Set(ByVal value As String)
            SLeadSource = value
        End Set
    End Property
    Public Property Manage() As String
        Get
            Return MarketManage
        End Get
        Set(ByVal value As String)
            MarketManage = value
        End Set
    End Property
#End Region
#Region "Methods"

    '' need method(s) to: 
    '' 1) gather existing information
    '' 2) check to see if any information has changed.
    '' 3) makes sure data types correspond
    '' 4) update the table
    '' 5) depending on how called, requery for current information
    '' 
    Public Sub RemoveLinkLabel()
        Try
            Dim ctrl As Control
            For Each ctrl In EditCustomerInfo.Controls
                If TypeOf ctrl Is GroupBox Then
                    If ctrl.Name = "gbContactInfo" Then
                        Dim gb As GroupBox = ctrl
                        Dim ctrl2 As Control '' look for link label now
                        For Each ctrl2 In gb.Controls
                            If TypeOf ctrl2 Is LinkLabel Then
                                Dim lnk As LinkLabel = ctrl2
                                If lnk.Name = "lnkEmail" Then
                                    gb.Controls.Remove(lnk)
                                    Dim txtEmail As New TextBox
                                    txtEmail.Location = New Point(413, 142)
                                    txtEmail.BorderStyle = BorderStyle.Fixed3D
                                    txtEmail.Size = New Size(169, 23)
                                    txtEmail.Name = "txtEditEmail"
                                    txtEmail.Text = Me.EmailAddress
                                    gb.Controls.Add(txtEmail)
                                    txtEmail.Focus()
                                    txtEmail.Select()
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "RemoveLinkLabel")
        End Try
    End Sub
    Public Sub RemoveTextBox()
        Try
            Dim ctrl As Control
            For Each ctrl In EditCustomerInfo.Controls
                If TypeOf ctrl Is GroupBox Then
                    If ctrl.Name = "gbContactInfo" Then
                        Dim gb As GroupBox = ctrl
                        Dim ctrl2 As Control '' look for link label now
                        For Each ctrl2 In gb.Controls
                            If TypeOf ctrl2 Is TextBox Then
                                Dim txt As TextBox = ctrl2
                                If txt.Name = "txtEditEmail" Then
                                    gb.Controls.Remove(txt)
                                    Dim lnkEmail As New LinkLabel
                                    lnkEmail.Location = New Point(413, 145)
                                    'lnkEmail.BorderStyle = BorderStyle.Fixed3D
                                    lnkEmail.Size = New Size(169, 23)
                                    lnkEmail.Name = "lnkEmail"
                                    lnkEmail.Text = Me.EmailAddress
                                    gb.Controls.Add(lnkEmail)
                                    lnkEmail.Enabled = False
                                    'txtEmail.Focus()
                                    'txtEmail.Select()
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm.ToString, "Front_End", "RemoveTextBox")
        End Try
    End Sub
    Function Get_First(ByVal name) As String
        If name.ToString.Contains(" ") = True Then
            Dim str = Split(name, " ")
            name = Trim(str(0).ToString)
        End If

  
        Return name


    End Function
    Function Get_Last(ByVal name) As String
        If name.ToString.Contains(" ") = True Then
            Dim str = Split(name, " ")
            name = Trim(str(1).ToString)
        End If
     


        Return name
    End Function
    Function Empty_cbo(ByVal sel_item) As String
        If sel_item = Nothing Then
            sel_item = ""
        End If
        Return sel_item
    End Function
    Public Sub New_Properties()
        Me.Contact1FirstName = Get_First(EditCustomerInfo.txtContact1.Text)
        Me.Contact1LastName = Get_Last(EditCustomerInfo.txtContact1.Text)
        Me.Contact2FirstName = Get_First(EditCustomerInfo.txtContact2.Text)
        Me.Contact2LastName = Get_Last(EditCustomerInfo.txtContact2.Text)
        Me.MainPhone = Me.DoLiteralsExist(EditCustomerInfo.txtHousePhone.Text)
        Me.AltPhone1 = Me.DoLiteralsExist(EditCustomerInfo.txtaltphone1.Text)
        Me.AltPhone2 = Me.DoLiteralsExist(EditCustomerInfo.txtaltphone2.Text)
        Me.Alt1Type = Empty_cbo(EditCustomerInfo.cboalt1type.SelectedItem)
        Me.Alt2Type = Empty_cbo(EditCustomerInfo.cboAlt2Type.SelectedItem)
        Me.StAddress = EditCustomerInfo.txtAddress.Text
        Me.STATE = EditCustomerInfo.txtState.Text
        Me.City = EditCustomerInfo.txtCity.Text
        Me.Zip = EditCustomerInfo.txtZip.Text
        Me.Contact1WorkHours = Empty_cbo(EditCustomerInfo.cboC1WorkHours.SelectedItem)
        Me.Contact2WorkHours = Empty_cbo(EditCustomerInfo.cboC2WorkHours.SelectedItem)
        Me.EmailAddress = EditCustomerInfo.lnkEmail.Text
        Me.YearsOwned = EditCustomerInfo.txtYrsOwned.Text
        Me.HomeValue = EditCustomerInfo.txtHomeValue.Text
        Me.YearBuilt = EditCustomerInfo.txtYrBuilt.Text
        'Me.AppointmentDate = Me.CorrectedApptDate(EditCustomerInfo.txtApptDate.Value)
        'Me.AppointmentTime = Me.CorrectedApptTime(EditCustomerInfo.txtApptTime.Value)
        Me.AppointmentDay = EditCustomerInfo.txtApptDay.Text
        Me.Product1 = Empty_cbo(EditCustomerInfo.cboProduct1.SelectedItem)
        Me.Product2 = Empty_cbo(EditCustomerInfo.cboProduct2.SelectedItem)
        Me.Product3 = Empty_cbo(EditCustomerInfo.cboProduct3.SelectedItem)
        Me.Product_Quantity = EditCustomerInfo.txtQty.Text
        Me.Product_Color = EditCustomerInfo.txtColor.Text
        Me.Special_Instruction = EditCustomerInfo.rtbSpecialInstructions.Text
        Me.Marketer_ = Empty_cbo(EditCustomerInfo.cboMarketer.SelectedItem)
        Me.Manage = Me.MarketManage
        Me.PriLS = Empty_cbo(EditCustomerInfo.cboPriLead.SelectedItem)
        Me.SecLS = Empty_cbo(EditCustomerInfo.cboSecLead.SelectedItem)

        Me.UpdateRecord(EditCustomerInfo.ID)
    End Sub
    Public Sub FeedProperties(ByVal ID As String)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("dbo.GetCommonBlockInfo", cnn)
            cmdGET.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            cmdGET.Parameters.Add(param1)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            While r1.Read
                Me.Contact1FirstName = r1.Item("Contact1FirstName")
                Me.Contact1LastName = r1.Item("Contact1LastName")
                Me.Contact2FirstName = r1.Item("Contact2FirstName")
                Me.Contact2LastName = r1.Item("Contact2LastName")
                Me.MainPhone = Me.DoLiteralsExist(r1.Item("HousePhone"))

                Me.AltPhone1 = Me.DoLiteralsExist(r1.Item("AltPhone1"))
                Me.Alt1Type = Me.DoLiteralsExist(r1.Item("Phone1Type"))
                Me.Alt2Type = r1.Item("Phone2Type")
                Me.AltPhone2 = r1.Item("AltPhone2")
                Me.StAddress = r1.Item("StAddress")
                Me.STATE = r1.Item("State")
                Me.City = r1.Item("City")
                Me.Zip = r1.Item("Zip")
                Me.Contact1WorkHours = r1.Item("Contact1WorkHours")
                Me.Contact2WorkHours = r1.Item("Contact2WorkHours")
                Me.EmailAddress = r1.Item("EmailAddress")
                Me.YearsOwned = r1.Item("YearsOwned")
                Me.HomeValue = r1.Item("HomeValue")
                Me.YearBuilt = r1.Item("HomeAge")
                Me.AppointmentDate = Me.CorrectedApptDate(r1.Item("ApptDate").ToString)
                Me.AppointmentTime = Me.CorrectedApptTime(r1.Item("ApptTime").ToString)
                Me.AppointmentDay = r1.Item("ApptDay")
                Me.Product1 = r1.Item("Product1")
                Me.Product2 = r1.Item("Product2")
                Me.Product3 = r1.Item("Product3")
                Me.Product_Quantity = r1.Item("ProductQTY")
                Me.Product_Color = r1.Item("Color")
                Me.Special_Instruction = r1.Item("SpecialInstruction")
                Me.Marketer_ = r1.Item("Marketer")
                Me.Manage = r1.Item("MarketingManager")
                Me.PriLS = r1.Item("PrimaryLeadSource")
                Me.SecLS = r1.Item("SecondaryLeadSource")
             

            End While
            r1.Close()
            cnn.Close()
            Me.PopulateForm()
            Me.Mp_verified(ID)
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal ID as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "FeedProperties")

        End Try
    End Sub '' populuates initial object with existing information 1)
    Public Sub PopulateForm()
        Try
            'Me.GetType()
            Me.GetWorkHours()
            Me.GetMarketers()
            Me.GetPrimaryLeadSource()
            Me.GetCities()
            Me.GetProducts()



            EditCustomerInfo.txtContact1.Text = Me.Contact1FirstName & " " & Me.Contact1LastName
            EditCustomerInfo.txtContact2.Text = Me.Contact2FirstName & " " & Me.Contact2LastName
            EditCustomerInfo.txtHousePhone.Text = Me.MainPhone
            EditCustomerInfo.txtaltphone1.Text = Me.AltPhone1
            EditCustomerInfo.txtaltphone2.Text = Me.AltPhone2
            EditCustomerInfo.cboalt1type.Text = Me.Alt1Type
            EditCustomerInfo.cboAlt2Type.Text = Me.Alt2Type
            EditCustomerInfo.OAddy = Me.StAddress
            EditCustomerInfo.txtAddress.Text = Me.StAddress
            EditCustomerInfo.OCity = Me.City
            EditCustomerInfo.txtCity.Text = Me.City
            EditCustomerInfo.OState = Me.STATE
            EditCustomerInfo.txtState.Text = Me.STATE
            EditCustomerInfo.OZip = Me.Zip
            EditCustomerInfo.txtZip.Text = Me.Zip


            EditCustomerInfo.cboC1WorkHours.Text = Me.Contact1WorkHours
            EditCustomerInfo.cboC2WorkHours.Text = Me.Contact2WorkHours
            EditCustomerInfo.lnkEmail.Text = Me.EmailAddress
            EditCustomerInfo.txtYrBuilt.Text = Me.YearBuilt
            EditCustomerInfo.txtHomeValue.Text = Me.HomeValue
            EditCustomerInfo.txtYrsOwned.Text = Me.YearsOwned
            'EditCustomerInfo.txtHomeValue.Text = Me.HomeValue
            EditCustomerInfo.txtApptDate.Text = Me.AppointmentDate
            EditCustomerInfo.txtApptTime.Text = Me.AppointmentTime
            EditCustomerInfo.txtApptDay.Text = Me.AppointmentDay

            Dim i '= EditCustomerInfo.cboP1.Items
            Dim exists As Boolean = False
            For Each i In EditCustomerInfo.cboProduct1.Items
                If Me.Product1 = i Then
                    exists = True
                End If
            Next
            If exists = False Then
                EditCustomerInfo.cboProduct1.Items.Add(Me.Product1)
                EditCustomerInfo.cboProduct1.SelectedItem = Me.Product1
            Else
                EditCustomerInfo.cboProduct1.SelectedItem = Me.Product1
            End If
            exists = False
            For Each i In EditCustomerInfo.cboProduct2.Items
                If Me.Product2 = i Then
                    exists = True
                End If
            Next
            If exists = False Then
                EditCustomerInfo.cboProduct2.Items.Add(Me.Product2)
                EditCustomerInfo.cboProduct2.SelectedItem = Me.Product2
            Else
                EditCustomerInfo.cboProduct2.SelectedItem = Me.Product2
            End If
            exists = False
            For Each i In EditCustomerInfo.cboProduct3.Items
                If Me.Product3 = i Then
                    exists = True
                End If
            Next
            If exists = False Then
                EditCustomerInfo.cboProduct3.Items.Add(Me.Product3)
                EditCustomerInfo.cboProduct3.SelectedItem = Me.Product3
            Else
                EditCustomerInfo.cboProduct3.SelectedItem = Me.Product3
            End If
            exists = False
            For Each i In EditCustomerInfo.cboPriLead.Items
                If Me.PriLS = i Then
                    exists = True
                End If
            Next
            If exists = False Then
                EditCustomerInfo.cboPriLead.Items.Add(Me.PriLS)
                EditCustomerInfo.cboPriLead.SelectedItem = Me.PriLS
            Else
                EditCustomerInfo.cboPriLead.SelectedItem = Me.PriLS
            End If
            exists = False
            For Each i In EditCustomerInfo.cboSecLead.Items
                If Me.SecLS = i Then
                    exists = True
                End If
            Next
            If exists = False Then
                EditCustomerInfo.cboSecLead.Items.Add(Me.SecLS)
                EditCustomerInfo.cboSecLead.SelectedItem = Me.SecLS
            Else
                EditCustomerInfo.cboSecLead.SelectedItem = Me.SecLS
            End If
            exists = False
            For Each i In EditCustomerInfo.cboMarketer.Items
                If Me.Marketer_ = i Then
                    exists = True
                End If
            Next
            If exists = False Then
                EditCustomerInfo.cboMarketer.Items.Add(Me.Marketer_)
                EditCustomerInfo.cboMarketer.SelectedItem = Me.Marketer_
            Else
                EditCustomerInfo.cboMarketer.SelectedItem = Me.Marketer_
            End If
            EditCustomerInfo.txtColor.Text = Me.Product_Color
            EditCustomerInfo.txtQty.Text = Me.Product_Quantity
            EditCustomerInfo.rtbSpecialInstructions.Text = Me.Special_Instruction
            EditCustomerInfo.cboMarketer.SelectedItem = Me.Marketer
            EditCustomerInfo.cboPriLead.SelectedItem = Me.PriLS
            Me.GetSLS(Me.PriLS)
            EditCustomerInfo.cboSecLead.SelectedItem = Me.SecLS


            'EditCustomerInfo.lnkEmail.Enabled = False
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "PopulateForm")

        End Try

    End Sub

#End Region
#Region "Instantiation"

#End Region
#Region "Functions"
    Private Function CorrectedApptTime(ByVal ApptTime As String)
        '' Example: "4/2/2008 12:00:00 AM"
        ''
        Try
            If ApptTime.ToString.Length <= 0 Then '' if there is no appttime return nothing.
                Me.CorrectedAppointmentTime = ""
                Return Trim(Me.CorrectedAppointmentTime)
                Exit Function
            End If
            Dim str = Split(ApptTime, " ", -1) '' should come back str(2) or 3 dimensional array
            Dim dt As String = Trim(str(0).ToString) '' date portion
            Dim AMPM As String = Trim(str(2).ToString) '' am/pm designator
            Dim time = Split(str(1), ":", -1) '' should return time(2) or 3 dimensional array
            Dim hour As String = Trim(time(0).ToString) '' hour
            Dim minute As String = Trim(time(1).ToString) '' minute
            Me.CorrectedAppointmentTime = hour & ":" & minute & " " & AMPM
            Return Trim(Me.CorrectedAppointmentTime)
        Catch ex As Exception
            Return Trim(Me.CorrectedAppointmentTime)
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal AppTime as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CorrectedApptTime")
        End Try


    End Function
    Private Function CorrectedApptDate(ByVal ApptDate As String)
        '' Example: "1/1/1900 11:00:00 AM"
        '' 
        Try
            If ApptDate.ToString.Length <= 0 Then '' if there is no appttime return nothing.
                Me.CorrectedAppointmentDate = ""
                Return Trim(Me.CorrectedAppointmentDate)
                Exit Function
            End If

            Dim str = Split(ApptDate, " ", -1) '' should come back str(2) or 3 dimensional array
            Dim dt As String = Trim(str(0).ToString) '' date portion
            Dim AMPM As String = Trim(str(2).ToString) '' am/pm designator
            Dim time = Split(str(1), ":", -1) '' should return time(2) or 3 dimensional array
            Dim hour As String = Trim(time(0).ToString) '' hour
            Dim minute As String = Trim(time(1).ToString) '' minute
            Me.CorrectedAppointmentDate = dt.ToString
            Return Trim(Me.CorrectedAppointmentDate)
            Return Me.CorrectedAppointmentDate
        Catch ex As Exception
            Return Trim(Me.CorrectedAppointmentDate)
            Return Me.CorrectedAppointmentDate
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CorrectedApptDate")
        End Try

    End Function
    'Private Function CorrectHValue(ByVal HouseValue As String)
    '    '' Example: 125.0000 = "$125,000"
    '    ''
    '    Try
    '        If HouseValue.ToString.Length <= 0 Then
    '            HouseValue = "0"
    '            Return Me.CorrectedHouseValue
    '            Exit Function
    '        End If

    '        Dim str = Split(HouseValue, ".", -1) 'returns 125|0000
    '        HouseValue = "$" & str(0) & ",000"
    '        'HouseValue = "$" & Format(HouseValue, "###,###.00")
    '        CorrectedHouseValue = HouseValue
    '        Return Me.CorrectedHouseValue
    '    Catch ex As Exception
    '        Return Me.CorrectedHouseValue
    '        Dim err As New ErrorLogFlatFile
    '        err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal HouseValue as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CorrectHValue")
    '    End Try

    'End Function
    'Private Function ConvertHouseValueToDecimal(ByVal HouseValue As String)
    '    '' example: $125,000 = "125.0000"
    '    ''
    '    Try
    '        If HouseValue.ToString.Length <= 0 Then
    '            HouseValue = "0"
    '            Return Me.CorrectedHouseValue
    '            Exit Function
    '        End If

    '        Dim str = Split(HouseValue, ",", -1) 'returns 125|000
    '        HouseValue = str(0).ToString 'returns '$125'
    '        HouseValue = Replace(str(0).ToString, "$", " ")
    '        HouseValue = Trim(HouseValue)
    '        CorrectedHouseValue = HouseValue
    '        Return Me.CorrectedHouseValue
    '    Catch ex As Exception
    '        Return Me.CorrectedHouseValue
    '        Dim err As New ErrorLogFlatFile
    '        err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal HouseValue as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "ConvertHouseValueToDecimal")

    '    End Try

    'End Function
    Public Sub UpdateRecord(ByVal ID As String)
        Try
            Dim cmdUP As SqlCommand = New SqlCommand("dbo.EditCustomerInformation", cnn)
            cmdUP.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            Dim param2 As SqlParameter = New SqlParameter("@C1FName", Me.Contact1FirstName)
            Dim param3 As SqlParameter = New SqlParameter("@C1LName", Me.Contact1LastName)
            Dim param4 As SqlParameter = New SqlParameter("@C2FName", Me.Contact2FirstName)
            Dim param5 As SqlParameter = New SqlParameter("@C2LName", Me.Contact2LastName)
            Dim param6 As SqlParameter = New SqlParameter("@HPhone", Me.MainPhone)
            Dim param7 As SqlParameter = New SqlParameter("@AltPhone1", Me.AltPhone1)
            Dim param8 As SqlParameter = New SqlParameter("@AltPhone2", Me.AltPhone2)
            Dim param9 As SqlParameter = New SqlParameter("@Phone1Type", Me.Alt1Type)
            Dim param10 As SqlParameter = New SqlParameter("@Phone2Type", Me.Alt2Type)
            Dim param11 As SqlParameter = New SqlParameter("@StAddress", Me.StAddress)
            Dim param12 As SqlParameter = New SqlParameter("@State", Me.STATE)
            Dim param13 As SqlParameter = New SqlParameter("@Zip", Me.Zip)
            Dim param14 As SqlParameter = New SqlParameter("@City", Me.City)
            Dim param15 As SqlParameter = New SqlParameter("@C1Work", Me.Contact1WorkHours)
            Dim param16 As SqlParameter = New SqlParameter("@C2Work", Me.Contact2WorkHours)
            Dim param17 As SqlParameter = New SqlParameter("@Email", Me.EmailAddress)
            Dim param18 As SqlParameter = New SqlParameter("@YearsOwned", Me.YearsOwned)

            Dim param19 As SqlParameter = New SqlParameter("@HomeValue", Me.HomeValue)
            Dim param20 As SqlParameter = New SqlParameter("@HomeAge", Me.YearBuilt)
            ' Dim param21 As SqlParameter = New SqlParameter("@Apptdate", Me.AppointmentDate)
            ' Dim param22 As SqlParameter = New SqlParameter("@ApptTime", Me.AppointmentTime)
            ' Dim param23 As SqlParameter = New SqlParameter("@ApptDay", Me.AppointmentDay)
            Dim param24 As SqlParameter = New SqlParameter("@Product1", Me.Product1)
            Dim param25 As SqlParameter = New SqlParameter("@Product2", Me.Product2)
            Dim param26 As SqlParameter = New SqlParameter("@Product3", Me.Product3)
            Dim param27 As SqlParameter = New SqlParameter("@ProductQTY", Me.Product_Quantity)
            Dim param28 As SqlParameter = New SqlParameter("@Color", Me.Product_Color)
            Dim param29 As SqlParameter = New SqlParameter("@SpecialInstruction", Me.Special_Instruction)
            Dim param30 As SqlParameter = New SqlParameter("@Marketer", Me.Marketer_)
            Dim param31 As SqlParameter = New SqlParameter("@PS", Me.PriLS)
            Dim param32 As SqlParameter = New SqlParameter("@SS", Me.SecLS)
            Dim param33 As SqlParameter = New SqlParameter("@MM", Me.Manage)


            cmdUP.Parameters.Add(param1)
            cmdUP.Parameters.Add(param2)
            cmdUP.Parameters.Add(param3)
            cmdUP.Parameters.Add(param4)
            cmdUP.Parameters.Add(param5)
            cmdUP.Parameters.Add(param6)
            cmdUP.Parameters.Add(param7)
            cmdUP.Parameters.Add(param8)
            cmdUP.Parameters.Add(param9)
            cmdUP.Parameters.Add(param10)
            cmdUP.Parameters.Add(param11)
            cmdUP.Parameters.Add(param12)
            cmdUP.Parameters.Add(param13)
            cmdUP.Parameters.Add(param14)
            cmdUP.Parameters.Add(param15)
            cmdUP.Parameters.Add(param16)
            cmdUP.Parameters.Add(param17)
            cmdUP.Parameters.Add(param18)
            cmdUP.Parameters.Add(param19)
            cmdUP.Parameters.Add(param20)
            ' cmdUP.Parameters.Add(param21)
            ' cmdUP.Parameters.Add(param22)
            ' cmdUP.Parameters.Add(param23)
            cmdUP.Parameters.Add(param24)
            cmdUP.Parameters.Add(param25)
            cmdUP.Parameters.Add(param26)
            cmdUP.Parameters.Add(param27)
            cmdUP.Parameters.Add(param28)
            cmdUP.Parameters.Add(param29)
            cmdUP.Parameters.Add(param30)
            cmdUP.Parameters.Add(param31)
            cmdUP.Parameters.Add(param32)
            cmdUP.Parameters.Add(param33)

            cmdUP.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            MsgBox("Problem update record: " & ID.ToString & ". Please try again.", MsgBoxStyle.Exclamation, "ERROR UPDATING")
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal ID As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "UpdateRecord")

        End Try
    End Sub
#End Region
    Public Sub GetMarketers()
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim dset_Marketers As Data.DataSet = New Data.DataSet("MARKETERS")
        Dim da_Marketers As SqlDataAdapter = New SqlDataAdapter



        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetMarketers", cnn)
            cmdGet.CommandType = CommandType.StoredProcedure
            cnn.Open()
            da_Marketers.SelectCommand = cmdGet
            da_Marketers.Fill(dset_Marketers, "MarketerPullList")
            Dim cnt As Integer = 0
            cnt = dset_Marketers.Tables(0).Rows.Count
            Select Case cnt
                Case Is <= 0
                    EditCustomerInfo.cboMarketer.Items.Clear()
                    EditCustomerInfo.cboMarketer.Items.Add("<Add New>")
                    EditCustomerInfo.cboMarketer.Items.Add("_____________________________________________")
                    EditCustomerInfo.cboMarketer.Items.Add("")
                    Exit Select
                Case Is >= 1
                    EditCustomerInfo.cboMarketer.Items.Clear()
                    EditCustomerInfo.cboMarketer.Items.Add("<Add New>")
                    EditCustomerInfo.cboMarketer.Items.Add("_____________________________________________")
                    EditCustomerInfo.cboMarketer.Items.Add("")
                    Dim b
                    For b = 0 To dset_Marketers.Tables(0).Rows.Count - 1
                        EditCustomerInfo.cboMarketer.Items.Add(dset_Marketers.Tables(0).Rows(b).Item(1) & " " & dset_Marketers.Tables(0).Rows(b).Item(2))
                    Next
                    Exit Select
            End Select
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.GetMarketers", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetMarketers")

        End Try

    End Sub
    'Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public dset_PriLS As Data.DataSet = New Data.DataSet("PLS")
    Public da_PRI As SqlDataAdapter = New SqlDataAdapter
    Public Sub GetPrimaryLeadSource()
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetPLS", cnn)
            cmdGet.CommandType = CommandType.StoredProcedure
            cnn.Open()
            da_PRI.SelectCommand = cmdGet
            da_PRI.Fill(dset_PriLS, "PrimaryLeadSource")
            Dim cnt As Integer = 0
            cnt = dset_PriLS.Tables(0).Rows.Count
            Select Case cnt
                Case Is <= 0
                    EditCustomerInfo.cboPriLead.Items.Clear()
                    EditCustomerInfo.cboPriLead.Items.Add("<Add New>")
                    EditCustomerInfo.cboPriLead.Items.Add("_____________________________________________")
                    EditCustomerInfo.cboPriLead.Items.Add("")
                    Exit Select
                Case Is >= 1
                    EditCustomerInfo.cboPriLead.Items.Clear()
                    EditCustomerInfo.cboPriLead.Items.Add("<Add New>")
                    EditCustomerInfo.cboPriLead.Items.Add("_____________________________________________")
                    EditCustomerInfo.cboPriLead.Items.Add("")
                    Dim b
                    For b = 0 To dset_PriLS.Tables(0).Rows.Count - 1
                        EditCustomerInfo.cboPriLead.Items.Add(dset_PriLS.Tables(0).Rows(b).Item(1))
                    Next
                    Exit Select
            End Select
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.GetPrimaryLeadSource", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetPrimaryLeadSource")

        End Try

    End Sub

    Public dset_SLS As Data.DataSet = New Data.DataSet("SLS")
    Public da_SLS As SqlDataAdapter = New SqlDataAdapter
    Public Sub GetSLS(ByVal PrimaryLS As String)
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetSLS", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@PRILS", PrimaryLS)
            cmdGet.Parameters.Add(param1)
            cnn.Open()
            cmdGet.CommandType = CommandType.StoredProcedure
            da_SLS.SelectCommand = cmdGet
            da_SLS.Fill(dset_SLS, "SecondaryLeadSource")
            Dim cnt As Integer = 0
            cnt = dset_SLS.Tables(0).Rows.Count
            Select Case cnt
                Case Is <= 0
                    EditCustomerInfo.cboSecLead.Items.Clear()
                    EditCustomerInfo.cboSecLead.Items.Add("<Add New>")
                    EditCustomerInfo.cboSecLead.Items.Add("_____________________________________________")
                    EditCustomerInfo.cboSecLead.Items.Add("")
                    Exit Select
                Case Is >= 1
                    EditCustomerInfo.cboSecLead.Items.Clear()
                    EditCustomerInfo.cboSecLead.Items.Add("<Add New>")
                    EditCustomerInfo.cboSecLead.Items.Add("_____________________________________________")
                    EditCustomerInfo.cboSecLead.Items.Add("")
                    Dim b As Integer = 0
                    For b = 0 To dset_SLS.Tables(0).Rows.Count - 1
                        EditCustomerInfo.cboSecLead.Items.Add(dset_SLS.Tables(0).Rows(b).Item(0))
                    Next
                    Exit Select
            End Select
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.GetSecondaryLeadSource", "ByVal PrimaryLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetSLS")
        End Try
    End Sub
    Public Sub GetMarketerLeadSources(ByVal FName As String, ByVal LName As String)
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetMarketerLeadSources", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@FName", FName)
            Dim param2 As SqlParameter = New SqlParameter("@LName", LName)
            cmdGet.Parameters.Add(param1)
            cmdGet.Parameters.Add(param2)
            cnn.Open()
            cmdGet.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader
            While r1.Read

                Me.MarketManage = r1.Item(2)
                EditCustomerInfo.cboPriLead.Text = ""
                EditCustomerInfo.cboPriLead.Text = r1.Item(0)
                EditCustomerInfo.cboSecLead.Text = ""
                EditCustomerInfo.cboSecLead.Text = r1.Item(1)

            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.GetMarketerLeadSources", "ByVal FName As String, ByVal LName As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetMarketerLeadSources")

        End Try

    End Sub
    Public Sub GetCities()
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetCities", cnn)
            cmdGet.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader
            EnterLead.txtCity.AutoCompleteCustomSource.Clear()
            While r1.Read
                EnterLead.txtCity.AutoCompleteCustomSource.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.GetCities", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetCities")
        End Try


    End Sub
    Public Sub GetWorkHours()
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetWorkHours", cnn)
            cnn.Open()
            cmdGet.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader
            EditCustomerInfo.cboC2WorkHours.Items.Clear()
            EditCustomerInfo.cboC1WorkHours.Items.Clear()
            EditCustomerInfo.cboC2WorkHours.Items.Add("<Add New>")
            EditCustomerInfo.cboC1WorkHours.Items.Add("<Add New>")
            EditCustomerInfo.cboC1WorkHours.Items.Add("_________________________________")
            EditCustomerInfo.cboC2WorkHours.Items.Add("_________________________________")
            EditCustomerInfo.cboC2WorkHours.Items.Add("")
            EditCustomerInfo.cboC1WorkHours.Items.Add("")

            While r1.Read
                EditCustomerInfo.cboC1WorkHours.Items.Add(r1.Item(0))
                EditCustomerInfo.cboC2WorkHours.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.GetWorkHours", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetWorkHours")
        End Try

    End Sub
    Public Sub GetProducts()
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetProducts", cnn)
            Dim r1 As SqlDataReader
            cnn.Open()
            r1 = cmdGet.ExecuteReader
            EditCustomerInfo.cboProduct1.Items.Clear()
            EditCustomerInfo.cboProduct2.Items.Clear()
            EditCustomerInfo.cboProduct3.Items.Clear()
            EditCustomerInfo.cboProduct1.Items.Add("<Add New>")
            EditCustomerInfo.cboProduct2.Items.Add("<Add New>")
            EditCustomerInfo.cboProduct3.Items.Add("<Add New>")
            EditCustomerInfo.cboProduct1.Items.Add("___________________________________________")
            EditCustomerInfo.cboProduct2.Items.Add("___________________________________________")
            EditCustomerInfo.cboProduct3.Items.Add("___________________________________________")
            EditCustomerInfo.cboProduct1.Items.Add("")
            EditCustomerInfo.cboProduct2.Items.Add("")
            EditCustomerInfo.cboProduct3.Items.Add("")
            While r1.Read
                EditCustomerInfo.cboProduct1.Items.Add(r1.Item(0))
                EditCustomerInfo.cboProduct2.Items.Add(r1.Item(0))
                EditCustomerInfo.cboProduct3.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.GetProducts", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetProducts")

        End Try

    End Sub
    Public Sub AddNewProduct(ByVal Product As String, ByVal ProdACRO As String, ByVal CBO As String)
        Try
            Dim cmdCNT As SqlCommand = New SqlCommand("SELECT count(ID) from iss.dbo.products where Product = @Product", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Product", Product)
            cmdCNT.Parameters.Add(param1)
            cnn.Open()
            Dim cnt As Integer = 0
            Dim r1 As SqlDataReader
            r1 = cmdCNT.ExecuteReader
            While r1.Read
                cnt = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
            Select Case cnt
                Case Is < 1
                    Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertProduct", cnn)
                    cmdINS.CommandType = CommandType.StoredProcedure
                    Dim param2 As SqlParameter = New SqlParameter("@Product", Product)
                    Dim param3 As SqlParameter = New SqlParameter("@ProdACRO", ProdACRO)
                    cnn.Open()
                    cmdINS.Parameters.Add(param2)
                    cmdINS.Parameters.Add(param3)
                    Dim r2 As SqlDataReader
                    r2 = cmdINS.ExecuteReader
                    r2.Close()
                    cnn.Close()

                    Me.GetProducts()
                    Select Case CBO
                        Case Is = "CBO1"
                            EditCustomerInfo.cboProduct1.SelectedItem = Product
                            Exit Select
                        Case Is = "CBO2"
                            EditCustomerInfo.cboProduct2.SelectedItem = Product
                            Exit Select
                        Case Is = "CBO3"
                            EditCustomerInfo.cboProduct3.SelectedItem = Product
                            Exit Select
                    End Select
                Case Is >= 1
                    MsgBox("Duplicate Product exists. Please enter a new one.", MsgBoxStyle.Exclamation)
                    Exit Sub
            End Select
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_CUSTOMER_INFORMATION.AddNewProduct", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "AddNewProduct")
        End Try

    End Sub
    Public Sub InsertMarketer(ByVal frm As Form)
        Try
    
            MarketerLeadSources.Frm = frm
            MarketerLeadSources.ShowInTaskbar = False
            MarketerLeadSources.ShowDialog()

        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD.InsertMarketer", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "InsertMarketer")
        End Try

    End Sub
    Public Sub WriteMarketerToTable(ByVal MFname As String, ByVal MLName As String, ByVal PLS As String, ByVal SLS As String, ByVal MarketingMan As String)
        Try
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertMarkter", cnn)
            cmdINS.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@Fname", MFname)
            Dim param2 As SqlParameter = New SqlParameter("@LName", MLName)
            Dim param3 As SqlParameter = New SqlParameter("@PriLeadSource", PLS)
            Dim param4 As SqlParameter = New SqlParameter("@SecLeadSource", SLS)
            Dim param5 As SqlParameter = New SqlParameter("@MarketingManager", MarketingMan)
            cmdINS.Parameters.Add(param1)
            cmdINS.Parameters.Add(param2)
            cmdINS.Parameters.Add(param3)
            cmdINS.Parameters.Add(param4)
            cmdINS.Parameters.Add(param5)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdINS.ExecuteReader
            r1.Close()
            cnn.Close()

            GetPrimaryLeadSource()

            GetMarketers()

            EditCustomerInfo.cboMarketer.SelectedItem = MFname & " " & MLName
            MarketerLeadSources.MFName = ""
            MarketerLeadSources.MLName = ""
            MarketerLeadSources.cboSLS.Text = ""
            MarketerLeadSources.cboPRILS.Text = ""
            MarketerLeadSources.cboMM.Text = ""
            MarketerLeadSources.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD.InsertMarketer", "ByVal MFname As String, ByVal MLName As String, ByVal PLS As String, ByVal SLS As String, ByVal MarketingMan As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "WriteMarketerToTable")

        End Try

    End Sub
    Public Function CapitalizeText(ByVal TextToCap As String)
        Try
            Dim Text
            Text = Mid(TextToCap, 1, 1)
            Text = Text.ToString.ToUpper
            TextToCap = Text & Mid(TextToCap, 2, TextToCap.Length)
            Return TextToCap
        Catch ex As Exception
            Return TextToCap
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD.InsertMarketer", "ByVal TextToCap as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CapitalizeText")

        End Try

    End Function
    Public Sub InsertNewPLS(ByVal PLS As String)
        Try
            Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.PrimaryLeadSource WHERE PrimaryLead = @PLS", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)

            PLS = CapitalizeText(PLS)
            cmdCNT.Parameters.Add(param1)
            Dim cnt As Integer = 0
            Select Case cnt
                Case Is >= 1
                    MsgBox("Duplicate Primary Lead Source Exists. Please enter a new one.", MsgBoxStyle.Exclamation, "ERROR")
                    Exit Sub
                Case Is < 1
                    Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertPLS", cnn)
                    Dim param2 As SqlParameter = New SqlParameter("@PLS", PLS)
                    cmdINS.Parameters.Add(param2)
                    cmdINS.CommandType = CommandType.StoredProcedure
                    cnn.Open()
                    Dim r1 As SqlDataReader
                    r1 = cmdINS.ExecuteReader
                    r1.Close()
                    cnn.Close()
                    'EnterLead.cboPriLead.SelectedItem = PLS

                    GetPrimaryLeadSource()
                    EditCustomerInfo.cboPriLead.SelectedItem = PLS
            End Select
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD.InsertPLS", "ByVal PLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertNewPLS")

        End Try

    End Sub
    Public Sub InsertWH(ByVal str As String, ByVal cbo As ComboBox)
        Try
            Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.WorkHours WHERE Hour = '" & str & "'", cnn)
            'Dim param1 As SqlParameter = New SqlParameter("@str", str)

            'cmdCNT.Parameters.Add(param1)

            cnn.Open()
            Dim cnt As Integer = 0
            Dim r1 As SqlDataReader
            r1 = cmdCNT.ExecuteReader
            While r1.Read
                cnt = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
            Select Case cnt
                Case Is >= 1
                    MsgBox("Duplicate Work Hours Exists.Please try again.", MsgBoxStyle.Exclamation, "ERROR")
                    cbo.Text = ""
                    Exit Sub
                Case Is < 1
                    Dim cmdINS As SqlCommand = New SqlCommand("Insert WorkHours (Hour) Values ('" & str & "')", cnn)

                    cmdINS.CommandType = CommandType.Text
                    cnn.Open()

                    cmdINS.ExecuteReader()
                    cnn.Close()
                    Me.GetWorkHours()

                    cbo.SelectedItem = str
                    Exit Select
            End Select
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD.InsertSLS", "ByVal PLS As String, ByVal SLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertSLS")

        End Try
    End Sub
    Public Sub InsertSLS(ByVal PLS As String, ByVal SLS As String)
        Try
            Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.SecondaryLeadSource WHERE SecondaryLead = @SLS and PrimaryLead = @PLS", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@SLS", SLS)
            Dim param2 As SqlParameter = New SqlParameter("@PLS", PLS)
            cmdCNT.Parameters.Add(param1)
            cmdCNT.Parameters.Add(param2)
            cnn.Open()
            Dim cnt As Integer = 0
            Dim r1 As SqlDataReader
            r1 = cmdCNT.ExecuteReader
            While r1.Read
                cnt = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
            Select Case cnt
                Case Is >= 1
                    MsgBox("Duplicate Seconadary Lead Source Exists.Please try again.", MsgBoxStyle.Exclamation, "ERROR")
                    Exit Sub
                Case Is < 1
                    Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertSLS", cnn)
                    Dim param3 As SqlParameter = New SqlParameter("@PrimaryLead", PLS)
                    Dim param4 As SqlParameter = New SqlParameter("@SecondaryLead", SLS)
                    cmdINS.Parameters.Add(param3)
                    cmdINS.Parameters.Add(param4)
                    cmdINS.CommandType = CommandType.StoredProcedure
                    cnn.Open()
                    Dim r2 As SqlDataReader
                    r2 = cmdINS.ExecuteReader
                    r2.Close()
                    cnn.Close()

                    GetSLS(PLS)
                    EditCustomerInfo.cboSecLead.SelectedItem = SLS
                    Exit Select
            End Select
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD.InsertSLS", "ByVal PLS As String, ByVal SLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertSLS")

        End Try

    End Sub

    Public Sub CheckCity(ByVal City As String, ByVal State As String)
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.CountCity", cnn)
            cmdGet.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@CTY", City)
            cmdGet.Parameters.Add(param1)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader
            Dim cnt As Integer = 0
            While r1.Read
                cnt = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
            Select Case cnt
                Case Is >= 1
                    Exit Sub
                Case Is <= 0
                    Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertNewCity", cnn)
                    cmdINS.CommandType = CommandType.StoredProcedure
                    Dim param2 As SqlParameter = New SqlParameter("@CITY", City)
                    Dim param3 As SqlParameter = New SqlParameter("@STATE", State)
                    cmdINS.Parameters.Add(param2)
                    cmdINS.Parameters.Add(param3)
                    cnn.Open()
                    Dim r2 As SqlDataReader
                    r2 = cmdINS.ExecuteReader(CommandBehavior.CloseConnection)
                    r2.Close()
                    cnn.Close()
                    EditCustomerInfo.txtCity.AutoCompleteCustomSource.Add(City)
                    Exit Select
            End Select
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD.IsCityInTable", "ByVal City As String, ByVal State As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "CheckCity")
        End Try

    End Sub
    Public Sub AutoFillState(ByVal city)
        Try
            Dim cmdGETState As SqlCommand = New SqlCommand("SELECT State from iss.dbo.CityPull where City = @city", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@city", city)
            cnn.Open()
            cmdGETState.Parameters.Add(param1)
            Dim r1 As SqlDataReader
            r1 = cmdGETState.ExecuteReader
            While r1.Read
                EditCustomerInfo.txtState.Text = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Private Function DoLiteralsExist(ByVal NumberToCheck As String) As String
        Dim len As Integer = 0
        len = NumberToCheck.Length
        '' length: 10 = No Number
        '' length: 14 = Has number
        Select Case len
            Case Is = 10
                NumberToCheck = ""
                Exit Select
            Case Is = 14
                NumberToCheck = NumberToCheck
                Exit Select
        End Select
        Return NumberToCheck
    End Function
    Public Sub Get_New_Appt(ByVal id As String)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("dbo.GetCommonBlockInfo", cnn)
            cmdGET.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            cmdGET.Parameters.Add(param1)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            While r1.Read
              
                Me.AppointmentDate = Me.CorrectedApptDate(r1.Item("ApptDate").ToString)
                Me.AppointmentTime = Me.CorrectedApptTime(r1.Item("ApptTime").ToString)
                Me.AppointmentDay = r1.Item("ApptDay")
              

            End While
            r1.Close()
            cnn.Close()

        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal ID as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "FeedProperties")

        End Try
    End Sub
    Public Sub Mp_verified(ByVal id As String)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("Select verified from verifiedaddress where leadnum = " & id, cnn)
            cmdGET.CommandType = CommandType.Text


            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            While r1.Read

                EditCustomerInfo.MP_Verified = r1.Item(0)


            End While
            r1.Close()
            cnn.Close()

        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal ID as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "FeedProperties")

        End Try
    End Sub
    Public Sub Update_MPVerified(ByVal id)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("Update verifiedaddress set verified = 'True'  where leadnum = " & id, cnn)
            cmdGET.CommandType = CommandType.Text
            cnn.Open()
            cmdGET.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal ID as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "FeedProperties")
        End Try
    End Sub
    Public Sub Update_MPVerified_False(ByVal id)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("Update verifiedaddress set verified = 'False'  where leadnum = " & id, cnn)
            cmdGET.CommandType = CommandType.Text
            cnn.Open()
            cmdGET.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("EDIT_CUSTOMER_INFORMATION", "ByVal ID as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "FeedProperties")
        End Try
    End Sub

End Class
