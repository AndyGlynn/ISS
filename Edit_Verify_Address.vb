
Imports System.Data
Imports MapPoint
Imports System.Data.Sql
Imports System
Imports System.Text
Imports System.Diagnostics.Process
Imports System.Data.SqlClient


Public Class Edit_Verify_Address
    Public oApp As MapPoint.Application
    Public oMap As MapPoint.Map
    Public oResults As MapPoint.FindResults
    Public oRoute As MapPoint.Route
    Public oWayPoints As MapPoint.Waypoints
    Public Loc As MapPoint.Location
    Public Loc2 As MapPoint.Location
    Public cntFilter1 As Integer = 0
    Public cntFilter2 As Integer = 0
    Public cntFilter3 As Integer = 0

    Public StAddress As String = ""
    Public City As String = ""
    Public State As String = ""
    Public Zip As String = ""

    Public NextFilterSwitch As Boolean = False

    Public Valid As Boolean
    Public CntValid As Integer = 0
    Public RunCnt As Integer = 1



    Public Morethan1 As Boolean
    Public LookForValid As Boolean = True
    Public ArValues As ArrayList
    Public MapPointVerified As Boolean
    Public CountOfLeads As Integer = 0

    Public CloseMethod As String = ""

    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.Cnn)


    Public Sub New(ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal RunCnt As Integer)
        Me.RunCnt = RunCnt

        '' Here is the first check to see if the city is in the table AFTER is has been pulled out of MapPoint
        ''
        ArValues = New ArrayList
        Filter(StAddress, City, State, Zip)

    End Sub
    ' '' General Process
    ' ''
    ' '' First: get count of found addresses through filter 1
    ' '' Second: if count of first is <=0 run second filter - 
    ' '' more ambiguous. Get count of second filter
    ' '' third: if count of second filter is <=0 then run third filter
    ' '' wich in turn is again more ambiguous.

    ' '' Constructor / Destructor Notes
    ' ''
    ' '' also need a disposing call here to free up mappoint
    ' '' because of its intermittent hangs. Process kill, object kill,
    ' '' or an actual IDisposable call/Method

    ' '' Additional Methods Needed
    ' ''
    ' '' need a method to verify that the address being pulled out is 
    ' '' a valid address IE: 123 Main Street and not W Main St

    ' '' Notes
    ' ''
    ' '' Need constructor arguments: StAddress, city, state, zip
    ' '' Need Filter 1 counter variable
    ' '' Need Filter 2 counter variable
    ' '' Need Filter 3 counter variable
    ' '' Need a switch to tell it to run next filter in line or not

    ' '' Need two checks here on the city for the AUTO-ADD to the city pull list.
    ' '' 1) check the city they are initially putting in. IE: check city in constructor object / sub
    ' '' 2) check the city after they have selected a new city from a MapPoint verified address.
    ' '' 3) because of a re-order issue, on the constructor object, the city must be verified through MapPoint
    ' ''    to make sure that it is a valid City/State name value pair.
    ' ''
    ' '' Two Additional Methods Added Here
    ' ''
    ' '' 1) A method to check to see if the city is in the list. //CheckCity(city,state)
    ' '' 2) A method just to verify on a duplicate lead whether or not the address has been verified already.
    ' ''    IE: A user hits "Add New Record" button on "Duplicate Lead" form // SingleVerify(LeadNumber)

    'Public Sub Filter1(ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String)
    '    '' must reorder check city method
    '    '' 
    '    Try
    '        oApp = New MapPoint.Application
    '        oMap = oApp.ActiveMap
    '        oResults = oMap.FindAddressResults(StAddress, City, , State, Zip, MapPoint.GeoCountry.geoCountryUnitedStates)
    '        cntFilter1 = oResults.Count
    '        If cntFilter1 <= 0 Then   ''counts results 
    '            NextFilterSwitch = True
    '            Me.MapPointVerified = False
    '            'MapPointVerified = False
    '        Else : NextFilterSwitch = False
    '            '' drop results to ArValues to be piped back out. 
    '            ''
    '            'MapPointVerified = True
    '            Me.MapPointVerified = True
    '            Dim c
    '            'Form1.lstAddresses.Items.Clear()
    '            Dim item As Integer = 0
    '            For Each c In oResults

    '                '' This loop is where AddressEnterLead comes into play
    '                ''
    '                '' need a count here of the results found, if more then one, pull up the choice form.
    '                '' if <=1 just roll on through.
    '                '' also, when choice is made on the front end, need a way to pipe it back to the rest of the 
    '                '' program. IE: what choice was made, and send it to ENTER_LEAD.InsertLead
    '                ''
    '                '' AddressEnterLead Considerations
    '                ''
    '                '' need overwrite ability on addresslist choice for mapped= true/false ?
    '                '' also, if "BACK" was hit, need a way to interrupt the logic so it doesnt complete the write,
    '                '' then go back to edit the lead information.
    '                '' 

    '                item += 1
    '                Loc1 = oResults(item)

    '                '' OR I run the check city logic here and run a check against ALL cities in the list
    '                '' that way, this action become less work intensive later on as the list will grow
    '                '' exponentially.
    '                '' 

    '                '' NEEDS to execute here while in this loop.
    '                '' MUST compensate for street addresses that are fuggled. IE: Ohio [OH], (Lucas) 


    '                If LookForValid = True Then
    '                    If LookForValidAddress(Loc1) = True Then
    '                        ArValues.Add(Loc1.Name.ToString)
    '                    End If
    '                Else
    '                    ArValues.Add(Loc1.Name.ToString)
    '                End If

    '                Dim CTY As String = ""
    '                Dim ST As String = ""

    '                Try
    '                    CTY = Loc1.StreetAddress.City.ToString
    '                    ST = Loc1.StreetAddress.Region.ToString
    '                Catch ex As Exception

    '                End Try

    '                'CheckCity(CTY, ST)
    '            Next

    '            '' Decision should go here AFTER results have been processed

    '            CountOfLeads = ArValues.Count
    '            Select Case CountOfLeads
    '                Case Is > 1
    '                    AddressEnterLead.lvAddresses.Items.Clear()
    '                    Dim lvI
    '                    For lvI = 1 To ArValues.Count
    '                        Dim lv As New ListViewItem
    '                        lv.Text = ArValues.Item(lvI - 1).ToString
    '                        AddressEnterLead.lvAddresses.Items.Add(lv)
    '                    Next
    '                    AddressEnterLead.ShowDialog()
    '                    '' after the decision has been made on the front end on what to do
    '                    '' depending on what the user chose, either keep processing or stop here and go back
    '                    '' and edit.
    '                    If AddressEnterLead.StopProcessing = True Then
    '                        Exit Sub
    '                    End If
    '                Case Is < 1
    '                    If NextFilterSwitch = True Then
    '                        Filter2(StAddress, City, State)
    '                    End If
    '                    'Form1.txtFilter1.Text = cntFilter1
    '                    oMap.Saved = True
    '                    oApp.Quit()
    '                    If NextFilterSwitch = False Then
    '                        KillProcess()
    '                    End If
    '                Case Is = 1
    '                    EnterLead.txtStAddy.Text = Loc1.StreetAddress.Street
    '                    EnterLead.txtCity.Text = Loc1.StreetAddress.City
    '                    EnterLead.txtState.Text = Loc1.StreetAddress.Region
    '                    EnterLead.txtZip.Text = Loc1.StreetAddress.PostalCode
    '                    EnterLead.MapPoint_PASSBACK = True
    '            End Select


    '            If CountOfLeads > 1 Then

    '            End If
    '            If CountOfLeads <= 1 Then

    '            End If
    '        End If
    '    Catch ex As Exception
    '        Dim err As New ErrorLogFlatFile
    '        err.WriteLog("VerifyAddress", "ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Mappoint", "Filter1")

    '    End Try

    'End Sub
    'Public Sub Filter2(ByVal StAddress As String, ByVal City As String, ByVal State As String)
    '    Try
    '        oApp = New MapPoint.Application
    '        oMap = oApp.ActiveMap
    '        oResults = oMap.FindAddressResults(StAddress, City, , State, , MapPoint.GeoCountry.geoCountryUnitedStates)
    '        cntFilter2 = oResults.Count
    '        If cntFilter2 <= 0 Then
    '            NextFilterSwitch = True
    '            'MapPointVerified = False
    '            Me.MapPointVerified = False

    '        Else : NextFilterSwitch = False
    '            'Dim colAddy As ListView = Form1.lstAddresses
    '            'Form1.lstAddresses.Columns.Clear()
    '            'colAddy.Width = 250
    '            'colAddy.Columns.Add("Address")
    '            ' colAddy.Columns(0).Width = 250
    '            'MapPointVerified = True
    '            Me.MapPointVerified = False
    '            Dim c
    '            'Form1.lstAddresses.Items.Clear()
    '            Dim item As Integer = 0
    '            For Each c In oResults
    '                item += 1
    '                Loc1 = oResults(item)
    '                If LookForValid = True Then
    '                    If LookForValidAddress(Loc1) = True Then
    '                        'Dim lv As New ListViewItem
    '                        'lv.Text = Loc1.Name.ToString
    '                        'Form1.lstAddresses.Items.Add(lv)
    '                        ArValues.Add(Loc1.Name.ToString)
    '                    End If
    '                Else
    '                    'Dim lv As New ListViewItem
    '                    'lv.Text = Loc1.Name.ToString
    '                    'Form1.lstAddresses.Items.Add(lv)
    '                    ArValues.Add(Loc1.Name.ToString)
    '                End If
    '            Next
    '        End If
    '        CountOfLeads = ArValues.Count
    '        If CountOfLeads > 1 Then
    '            AddressEnterLead.lvAddresses.Items.Clear()
    '            Dim lvI
    '            For lvI = 1 To ArValues.Count
    '                Dim lv As New ListViewItem
    '                lv.Text = ArValues.Item(lvI - 1).ToString
    '                AddressEnterLead.lvAddresses.Items.Add(lv)
    '            Next
    '            AddressEnterLead.ShowDialog()
    '            KillProcess()
    '            '' after the decision has been made on the front end on what to do
    '            '' depending on what the user chose, either keep processing or stop here and go back
    '            '' and edit.
    '            If AddressEnterLead.StopProcessing = True Then
    '                Exit Sub
    '            End If
    '        End If
    '        If NextFilterSwitch = True Then
    '            Filter3(StAddress, City)
    '        End If
    '        'Form1.txtFilter2.Text = cntFilter2
    '        oMap.Saved = True
    '        oApp.Quit()
    '        If NextFilterSwitch = False Then
    '            KillProcess()
    '        End If
    '    Catch ex As Exception
    '        Dim err As New ErrorLogFlatFile
    '        err.WriteLog("VerifyAddress", "ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Mappoint", "Filter2")

    '    End Try

    'End Sub
    'Public Sub Filter3(ByVal StAddress As String, ByVal City As String)
    '    Try
    '        oApp = New MapPoint.Application
    '        oMap = oApp.ActiveMap
    '        oResults = oMap.FindAddressResults(StAddress, , , State, , MapPoint.GeoCountry.geoCountryUnitedStates)
    '        cntFilter3 = oResults.Count
    '        'Form1.txtFilter3.Text = cntFilter3
    '        'Dim colAddy As ListView = Form1.lstAddresses
    '        'Form1.lstAddresses.Columns.Clear()
    '        'colAddy.Width = 250
    '        'colAddy.Columns.Add("Address")
    '        'colAddy.Columns(0).Width = 250
    '        'MapPointVerified = False
    '        Me.MapPointVerified = False
    '        Dim c
    '        'Form1.lstAddresses.Items.Clear()
    '        Dim item As Integer = 0
    '        For Each c In oResults
    '            item += 1
    '            Loc1 = oResults(item)
    '            If LookForValid = True Then
    '                If LookForValidAddress(Loc1) = True Then
    '                    'Dim lv As New ListViewItem
    '                    'lv.Text = Loc1.Name.ToString
    '                    'Form1.lstAddresses.Items.Add(lv)
    '                    ArValues.Add(Loc1.Name.ToString)
    '                End If
    '            Else
    '                'Dim lv As New ListViewItem
    '                'lv.Text = Loc1.Name.ToString
    '                'Form1.lstAddresses.Items.Add(lv)
    '                ArValues.Add(Loc1.Name.ToString)
    '            End If
    '        Next
    '        CountOfLeads = ArValues.Count
    '        If CountOfLeads > 1 Then
    '            AddressEnterLead.lvAddresses.Items.Clear()
    '            Dim lvI
    '            For lvI = 1 To ArValues.Count
    '                Dim lv As New ListViewItem
    '                lv.Text = ArValues.Item(lvI - 1).ToString
    '                AddressEnterLead.lvAddresses.Items.Add(lv)
    '            Next
    '            AddressEnterLead.ShowDialog()
    '            KillProcess()
    '            '' after the decision has been made on the front end on what to do
    '            '' depending on what the user chose, either keep processing or stop here and go back
    '            '' and edit.
    '            If AddressEnterLead.StopProcessing = True Then
    '                Exit Sub
    '            End If
    '        End If
    '        oMap.Saved = True
    '        oApp.Quit()
    '        KillProcess()
    '    Catch ex As Exception
    '        Dim err As New ErrorLogFlatFile
    '        err.WriteLog("VerifyAddress", "ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Mappoint", "Filter1")
    '    End Try
    'End Sub


    '' need method here to look to make sure that address pulled out
    '' is an actual valid address format

    Public Function LookForValidAddress(ByVal Loc As MapPoint.Location)
        '' conditions: No addresses like "W Main st", "Ohio [OH] (Ohio)"
        ''             MUST begin with Numbers
        Valid = False ''' 
        Try
            Dim Name As String = Loc.Name
            Dim ch As Char
            Dim cnt As Integer = 0
            For Each ch In Name
                If ch = "," Then
                    cnt += 1
                End If
            Next
            '' if there aren't enough commas, we dont want this address
            If cnt < 2 Then
                Valid = False
                Exit Function
            End If
            '' there are too many commas, we dont want this address
            If cnt > 2 Then
                Valid = False
                Exit Function
            End If
            '' if there exactly two commas, this is the address we want
            If cnt = 2 Then
                Dim STA As String = Loc.StreetAddress.Street
                Dim CTY As String = Loc.StreetAddress.City
                Dim ST As String = Loc.StreetAddress.Region
                Dim ZP As String = Loc.StreetAddress.PostalCode
                Dim LOCNAME = Split(Loc.Name, ",", 3)
                STA = LOCNAME(0)
                CTY = LOCNAME(1)
                Dim STATEZIP = Split(Trim(LOCNAME(2)), " ", 2)
                ST = STATEZIP(0)
                ZP = STATEZIP(1)
                If STA.Contains(" ") = False Then
                    Valid = False
                    Exit Function
                End If
                '' now we need to check the first part of the array
                '' to make sure it has a numeric street value
                '' otherwise it is garbage to us
                Dim i As Integer
                Dim c As Char
                For Each c In STA
                    i += 1
                    If i >= 4 Or c = " " Then
                        Exit For
                    End If
                    Select Case c
                        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                            Valid = True
                            'CntValid += 1
                        Case Else
                            Valid = False
                    End Select
                Next
            End If
            Return Valid
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("VerifyAddress", "ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Mappoint", "LookForValidAddress")

        End Try

    End Function
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
            err.WriteLog("VerifyAddress", "ByVal City As String, ByVal State As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Mappoint", "CheckCity")

        End Try

    End Sub
    Public Property MapPointVerifiedIt() As Boolean
        Get
            Return MapPointVerified
        End Get
        Set(ByVal value As Boolean)
            MapPointVerified = value
        End Set
    End Property
    Public Sub Filter(ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String)
        Try
            If Me.RunCnt = 4 Then
                Exit Sub
            End If


            oApp = STATIC_VARIABLES.oApp
            oMap = oApp.ActiveMap
            oResults = oMap.FindAddressResults(StAddress, City, , State, Zip, MapPoint.GeoCountry.geoCountryUnitedStates)
            Dim c
            Dim item As Integer = 0
            For Each c In oResults
                item += 1
                Me.Loc = oResults.Item(item)
                Me.LookForValidAddress(Me.Loc)
                If Me.Valid = True Then
                    Me.CntValid += 1
                    Me.ArValues.Add(Loc.Name)
                    Me.Loc2 = Me.Loc
                End If

            Next

            Select Case Me.ArValues.Count  ''count of actual valid addresses
                Case Is <= 0
                    Select Case Me.RunCnt
                        Case Is = 1

                            Dim x = New Edit_Verify_Address(StAddress, City, State, "", 2)

                            Exit Sub
                        Case Is = 2

                            Dim x = New Edit_Verify_Address(StAddress, City, "", "", 3)

                            Exit Sub
                        Case Is = 3
                            Dim response As Integer
                            response = MsgBox("Map Point was unable to find a valid address... Would you like to edit the address?" & _
                           " (By Selecting No your original entry will be used and the address will remain Un-Verified)", MsgBoxStyle.YesNo, "Map Point Cannot Verify Address")

                            If response = 6 Then

                                EditCustomerInfo.Focus()
                                EditCustomerInfo.txtAddress.SelectAll()
                                EditCustomerInfo.txtAddress.Focus()

                            ElseIf response = 7 Then
                                EditCustomerInfo.MP_Verified = False
                            End If

                            Exit Sub


                    End Select
                Case Is > 1
                    AddressEnterLead.lvAddresses.Items.Clear()
                    Dim lvI
                    For lvI = 1 To ArValues.Count
                        Dim lv As New ListViewItem
                        lv.Text = ArValues.Item(lvI - 1).ToString
                        AddressEnterLead.lvAddresses.Items.Add(lv)

                    Next
                    AddressEnterLead.frm = EditCustomerInfo
                    AddressEnterLead.StartPosition = FormStartPosition.CenterScreen
                    AddressEnterLead.ShowInTaskbar = False
                    AddressEnterLead.ShowDialog()
                    Exit Sub
                Case Is = 1
                    Me.LookForValidAddress(Me.Loc2)
                    If Me.Valid = True And Me.CntValid = 1 Then
                        EditCustomerInfo.txtAddress.Text = Me.Loc2.StreetAddress.Street
                        EditCustomerInfo.OAddy = Me.Loc2.StreetAddress.Street
                        EditCustomerInfo.OCity = Me.Loc2.StreetAddress.City
                        EditCustomerInfo.txtCity.Text = Me.Loc2.StreetAddress.City
                        EditCustomerInfo.txtState.Text = Me.Loc2.StreetAddress.Region
                        EditCustomerInfo.OState = Me.Loc2.StreetAddress.Region
                        EditCustomerInfo.txtZip.Text = Me.Loc2.StreetAddress.PostalCode
                        EditCustomerInfo.OZip = Me.Loc2.StreetAddress.PostalCode
                        EditCustomerInfo.MP_Verified = True
                        EditCustomerInfo.pctVerified.Visible = False
                        EditCustomerInfo.btnMap.Enabled = False
                        Dim z As New EDIT_CUSTOMER_INFORMATION
                        z.Update_MPVerified(EditCustomerInfo.ID)
                        Me.CheckCity(Me.Loc2.StreetAddress.City, Me.Loc2.StreetAddress.Region)
                        Exit Select
                    End If
            End Select
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("VerifyAddress", "ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Mappoint", "Filter1")
        End Try
    End Sub
End Class



