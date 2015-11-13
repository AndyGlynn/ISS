Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Public Class ManageAlerts
    Public ArLN As New ArrayList
    Public ArUID As New ArrayList
    Public ArAlertTime As New ArrayList
    Public ArNotes As New ArrayList
    Public arExecDate As New ArrayList
    Public Change As Boolean = False
    Public Delete As Boolean = False
    Public INDEX_OF_ARRAYLIST As Integer = 0
    Public MAX_INDEX_OF_ARRAYLIST As Integer = 0
    Public Reload As Boolean = False
    Public dset_Alerts As Data.DataSet = New Data.DataSet("ALERTS")
    Public lvct As Integer = 0
    Public i As ListViewItem

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChangeTime.Click
        'If Me.Size = New Size(298, 290) Then
        Me.Change = True
        Me.Button2_Click(Nothing, Nothing)
        '    Me.Size = New Size(298, 335)
        'ElseIf Me.Size = New Size(298, 335) Then
        '    Me.Size = New Size(298, 290)
        'End If
        'Me.Size = New Size(298, 335)
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        '' need a mechanism here to just change the text of the lnkID 
        '' the rest will be taken care of by the text changed event of the lnkID.....
        '' design time.
        '' 
        '' sloppy mechanism design, but will see if it works....
        'If Me.lvExpiredAlerts.Items.Count = 0 Then
        '    Exit Sub
        'End If
        If Change = True Then
            Dim str = Split(Me.dtpDate.Value, " ", 2)
            Dim time As String = "1/01/1900 " & Me.dtpTime.Value.ToShortTimeString
            Dim excdate As String = Me.dtpDate.Value.ToShortDateString & " 00:00:00"
            'MsgBox(excdate & "    " & time)
            UpdateAlertTime(Me.lnkID.Tag, time, excdate)
        ElseIf Delete = True Then
            Me.DeleteAlert(Me.lnkID.Tag)
        ElseIf Change = False And Delete = False Then
            If (INDEX_OF_ARRAYLIST + 1) > MAX_INDEX_OF_ARRAYLIST Then
                '' release resources
                '' 
                Me.dset_Alerts.Clear()
                Me.dset_Alerts = Nothing
                Me.ErrorProvider1.Clear()
                Me.ArAlertTime.Clear()
                Me.ArLN.Clear()
                Me.ArUID.Clear()
                Me.INDEX_OF_ARRAYLIST = 0
                Me.MAX_INDEX_OF_ARRAYLIST = 0
                Me.btnNext.Text = "Next"
                Me.Close()
            End If
            If (INDEX_OF_ARRAYLIST + 1) <= MAX_INDEX_OF_ARRAYLIST Then
                If (INDEX_OF_ARRAYLIST + 1) = MAX_INDEX_OF_ARRAYLIST Then
                    Me.btnNext.Text = "Done"
                ElseIf (INDEX_OF_ARRAYLIST + 1) < MAX_INDEX_OF_ARRAYLIST Then
                    Me.btnNext.Text = "Next"
                End If
                Me.lnkID.Tag = ArUID.Item(INDEX_OF_ARRAYLIST + 1)
                Me.lnkID.Text = ArLN.Item(INDEX_OF_ARRAYLIST + 1)
                Dim tm = Split(ArAlertTime.Item(INDEX_OF_ARRAYLIST + 1), " ", 2)
                Dim tim As String = tm(0)
                Dim peices = Split(tim, ":", 3)
                Dim ampm As String = tm(1)
                Dim convertedTime As String = peices(0) & ":" & peices(1) & " " & ampm
                Me.lblTime.Text = convertedTime
                Me.lblNotes.Text = ArNotes.Item(INDEX_OF_ARRAYLIST + 1)
                Dim ExD = arExecDate.Item(INDEX_OF_ARRAYLIST + 1)
                Me.lblDate.Text = ExD
                If Me.Change = True Then
                    '' now check to make sure the new time is at least a minute ahead.
                    '' 
                End If
                INDEX_OF_ARRAYLIST += 1
                ' Me.Size = New Size(298, 290)
                Me.Change = False
            End If
        End If
     
        If Me.lvExpiredAlerts.Items.Count <> 0 Then
            Me.lvExpiredAlerts.Items(0).Remove()
            Me.lvExpiredAlerts.Refresh()
        End If


        If Me.lvExpiredAlerts.Items.Count <> 0 Then
            Me.lvExpiredAlerts.TopItem.Selected = True
        End If


    End Sub

    Private Sub ManageAlerts_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Main.tmrAlerts.Start()
    End Sub


   




    Private Sub ManageAlerts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Me.SuspendLayout()
        Me.MdiParent = Main
        'Main.Refresh()
        'Me.Timer1.Start()
        Me.lvExpiredAlerts.Enabled = False

        '' populate the dataset to the form to be dealt with
        ''
        'Me.lblNotes.Size = New Size(250, 13)
        Me.ErrorProvider1.Clear()


        Dim r As Data.DataRow
        'If Reload = True Then
        '    Me.Reload = False
        '    Me.INDEX_OF_ARRAYLIST = 0
        '    Me.MAX_INDEX_OF_ARRAYLIST = ArUID.Count - 1
        '    Me.lnkID.Tag = ArUID.Item(0).ToString
        '    Me.lnkID.Text = ArLN.Item(0).ToString
        '    Me.lblNotes.Text = ArNotes.Item(0).ToString
        '    Dim tm = Split(ArAlertTime.Item(INDEX_OF_ARRAYLIST + 1), " ", 2)
        '    Dim tim As String = tm(0)
        '    Dim peices = Split(tim, ":", 3)
        '    Dim ampm As String = tm(1)
        '    Dim convertedTime As String = peices(0) & ":" & peices(1) & " " & ampm
        '    Me.lblTime.Text = convertedTime
        '    Exit Sub
        'End If
        For Each r In Me.dset_Alerts.Tables(0).Rows
            ArUID.Add(r.Item("ID").ToString)
            ArNotes.Add(r.Item("Notes").ToString)
            ArLN.Add(r.Item("LeadNUM").ToString)

            Dim tm = Split(r.Item("AlertTime").ToString, " ", 2)
            ArAlertTime.Add(tm(0).ToString & " " & tm(1).ToString)
            Dim ED = Split(r.Item("ExecutionDate"), " ", 2)
            arExecDate.Add(ED(0).ToString)
            Dim tm2 = Split(r.Item("AlertTime").ToString, " ", 2)
            Dim tim2 As String = tm2(0)
            Dim peices2 = Split(tim2, ":", 3)
            Dim ampm2 As String = tm2(1)
            Dim convertedTime2 As String = peices2(0) & ":" & peices2(1) & " " & ampm2
            Dim lv As New ListViewItem
            lv.Text = ED(0).ToString & " " & convertedTime2
            lv.SubItems.Add(r.Item("LeadNum"))
            lv.SubItems.Add(r.Item("CustFName") & " " & r.Item("CustLName"))
            lv.SubItems.Add(r.Item("Phone"))
            Me.lvExpiredAlerts.Items.Add(lv)
            lvct += 1

        Next
        Me.INDEX_OF_ARRAYLIST = 0
        Me.MAX_INDEX_OF_ARRAYLIST = ArUID.Count - 1
        Me.lnkID.Tag = ArUID.Item(0).ToString
        Me.lnkID.Text = ArLN.Item(0).ToString
        Me.lblNotes.Text = ArNotes.Item(0).ToString
        Dim tm1 = Split(ArAlertTime.Item(0), " ", 2) ''INDEX_OF_ARRAYLIST + 1
        Dim ExD = arExecDate.Item(0) ''INDEX_OF_ARRAYLIST + 1
        Dim tim1 As String = tm1(0)
        Dim peices1 = Split(tim1, ":", 3)
        Dim ampm1 As String = tm1(1)
        Dim convertedTime1 As String = peices1(0) & ":" & peices1(1) & " " & ampm1
        Me.lblTime.Text = convertedTime1
        Me.lblDate.Text = ExD.ToString
        If Me.lvct = 1 Then
            Me.btnNext.Text = "Done"
        End If
        Me.lvExpiredAlerts.TopItem.Selected = True
        Me.PerformLayout()
        Main.PerformLayout()


    End Sub



  
    Private Sub GetINFO(ByVal ID As String)
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT Contact1FirstName,Contact1LastName,Contact2FirstName,Contact2LastName,Staddress,city,state,zip,housephone,altphone1,altphone2,phone1type,phone2type from iss.dbo.enterlead where ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdGET.Parameters.Add(param1)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader
        While r1.Read
            Me.lblCustomerNFO.Text = r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName") & " and " & r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName") & vbCrLf & r1.Item("Staddress") & vbCrLf & r1.Item("City") & ", " & r1.Item("State") & ", " & r1.Item("Zip") & _
                 vbCrLf & r1.Item("HousePhone") & vbCrLf & r1.Item("AltPhone1") & "     " & r1.Item("Phone1Type") & vbCrLf & r1.Item("AltPhone2") & "     " & r1.Item("Phone2Type")
        End While
        r1.Close()
        cnn.Close()
    End Sub
    Private Sub DeleteAlert(ByVal ID As String)
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdDEL As SqlCommand = New SqlCommand("DELETE Iss.dbo.alerttable where ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdDEL.Parameters.Add(param1)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdDEL.ExecuteReader
        r1.Close()
        cnn.Close()
        If (INDEX_OF_ARRAYLIST + 1) > MAX_INDEX_OF_ARRAYLIST Then
            '' release resources
            '' 
            Me.dset_Alerts.Clear()
            Me.dset_Alerts = Nothing
            Me.ErrorProvider1.Clear()
            Me.ArAlertTime.Clear()
            Me.ArLN.Clear()
            Me.ArUID.Clear()
            Me.INDEX_OF_ARRAYLIST = 0
            Me.MAX_INDEX_OF_ARRAYLIST = 0
            Me.Close()
        End If
        If (INDEX_OF_ARRAYLIST + 1) <= MAX_INDEX_OF_ARRAYLIST Then
            Me.lnkID.Tag = ArUID.Item(INDEX_OF_ARRAYLIST + 1)
            Me.lnkID.Text = ArLN.Item(INDEX_OF_ARRAYLIST + 1)
            Dim tm = Split(ArAlertTime.Item(INDEX_OF_ARRAYLIST + 1), " ", 2)
            Dim tim As String = tm(0)
            Dim peices = Split(tim, ":", 3)
            Dim ampm As String = tm(1)
            Dim convertedTime As String = peices(0) & ":" & peices(1) & " " & ampm
            Me.lblTime.Text = convertedTime
            Dim ExD = arExecDate.Item(INDEX_OF_ARRAYLIST + 1)
            If Me.Change = True Then
                '' now check to make sure the new time is at least a minute ahead.
                '' need date change 
            End If
            If (INDEX_OF_ARRAYLIST + 1) = MAX_INDEX_OF_ARRAYLIST Then
                Me.btnNext.Text = "Done"
            ElseIf (INDEX_OF_ARRAYLIST + 1) < MAX_INDEX_OF_ARRAYLIST Then
                Me.btnNext.Text = "Next"
            End If
            INDEX_OF_ARRAYLIST += 1
            'Me.Size = New Size(298, 290)
            Me.Change = False
            Me.Delete = False
      
        End If
        '' after the alert is deleted, advance to the next id in line.
        '' if there is no other record after the record deleted, close the manage alerts form and release the 
        '' all resources, ie clear arraylists, clear datasets, blah blah blah.....
        '' 
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
    

       
        Me.Delete = True
        Me.Button2_Click(Nothing, Nothing)
        '' advance a record, or if you cannot advance. close form and release resources.
        '' blah blah blah....
    End Sub
    Private Sub UpdateAlertTime(ByVal ID As String, ByVal AlertTime As String, ByVal ExecutionDate As String)
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE iss.dbo.alerttable " _
        & " SET AlertTime = @TIME, " _
        & "     ExecutionDate = @ED, " _
        & "     Completed = 0" _
        & " WHERE ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@TIME", AlertTime)
        Dim param2 As SqlParameter = New SqlParameter("@ED", ExecutionDate)
        Dim param3 As SqlParameter = New SqlParameter("@ID", ID)
        cmdUP.Parameters.Add(param1)
        cmdUP.Parameters.Add(param2)
        cmdUP.Parameters.Add(param3)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdUP.ExecuteReader
        r1.Close()
        cnn.Close()
        If (INDEX_OF_ARRAYLIST + 1) > MAX_INDEX_OF_ARRAYLIST Then
            '' release resources
            '' 
            Me.dset_Alerts.Clear()
            Me.dset_Alerts = Nothing
            Me.ErrorProvider1.Clear()
            Me.ArAlertTime.Clear()
            Me.ArLN.Clear()
            Me.ArUID.Clear()
            Me.INDEX_OF_ARRAYLIST = 0
            Me.MAX_INDEX_OF_ARRAYLIST = 0
            Me.Close()
        End If
        If (INDEX_OF_ARRAYLIST + 1) <= MAX_INDEX_OF_ARRAYLIST Then
            Me.lnkID.Tag = ArUID.Item(INDEX_OF_ARRAYLIST + 1)
            Me.lnkID.Text = ArLN.Item(INDEX_OF_ARRAYLIST + 1)
            Dim tm = Split(ArAlertTime.Item(INDEX_OF_ARRAYLIST + 1), " ", 2)
            Dim tim As String = tm(0)
            Dim peices = Split(tim, ":", 3)
            Dim ampm As String = tm(1)
            Dim convertedTime As String = peices(0) & ":" & peices(1) & " " & ampm
            Me.lblTime.Text = convertedTime
            Dim ExD = arExecDate.Item(INDEX_OF_ARRAYLIST + 1)
            If Me.Change = True Then
                '' now check to make sure the new time is at least a minute ahead.
                '' 
            End If
            If (INDEX_OF_ARRAYLIST + 1) = MAX_INDEX_OF_ARRAYLIST Then
                Me.btnNext.Text = "Done"
            ElseIf (INDEX_OF_ARRAYLIST + 1) < MAX_INDEX_OF_ARRAYLIST Then
                Me.btnNext.Text = "Next"
            End If
            INDEX_OF_ARRAYLIST += 1
            'Me.Size = New Size(298, 290)
            Me.Change = False
         
        End If
        '' after the alert is deleted, advance to the next id in line.
        '' if there is no other record after the record deleted, close the manage alerts form and release the 
        '' all resources, ie clear arraylists, clear datasets, blah blah blah.....
        '' 
    End Sub

    Private Sub lnkID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkID.Click
        Dim found As Integer = 0
        Dim open As Boolean = False
        ' MsgBox(STATIC_VARIABLES.ActiveChild.Name.ToString)
        If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
            open = True
            Select Case STATIC_VARIABLES.ActiveChild.Name

                Case Is = "Confirming"
                    Dim tab As String = Confirming.Tab
                    'If tab = "Confirm" Then
                    Dim foundConfirming As Integer = 0
                    For Each Me.i In Confirming.lvConfirming.Items
                        If Me.i.Text = Me.lnkID.Text Then

                            Me.i.Selected = True
                            found += 1
                            foundConfirming += 1
                            If Confirming.TabControl1.SelectedIndex <> 0 Then
                                Confirming.TabControl1.SelectedIndex = 0
                            End If
                            Confirming.Focus()
                            Exit Sub


                            'Exit Sub
                        End If
                    Next
                    'End If
                    If Confirming.lvConfirming.SelectedItems.Count <> 0 Then
                        If Confirming.lvConfirming.SelectedItems(0).Text <> Me.lnkID.Text Then
                            For Each Me.i In Confirming.lvSales.Items
                                If Me.i.Text = Me.lnkID.Text Then
                                    Me.i.Selected = True
                                    found += 1
                                    foundConfirming += 1
                                    If Confirming.TabControl1.SelectedIndex <> 1 Then
                                        Confirming.TabControl1.SelectedIndex = 1
                                    End If
                                    Confirming.Focus()
                                    Exit Sub

                                    'Exit Sub
                                End If
                            Next
                        End If
                    End If

                    If foundConfirming = 0 Then
                        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                        Dim cmdGET As SqlCommand = New SqlCommand("dbo.AlertConfirm", cnn)

                        Dim param1 As SqlParameter = New SqlParameter("@ID", Me.lnkID.Text)
                        cmdGET.Parameters.Add(param1)
                        cmdGET.CommandType = CommandType.StoredProcedure
                        cnn.Open()
                        Dim r1 As SqlDataReader

                        r1 = cmdGET.ExecuteReader
                        While r1.Read
                            'MsgBox(r1.Item(0))
                            Dim was As Date = Confirming.dpConfirming.Value
                            Confirming.dpConfirming.Value = r1.Item(0)
                            'Exit Sub '' remove 
                            'MsgBox(Confirming.lvConfirming.Items.Count)
                            For Each Me.i In Confirming.lvConfirming.Items
                                'MsgBox(Confirming.lvConfirming.SelectedItems(0).Text)
                                If Me.i.Text = Me.lnkID.Text Then

                                    Me.i.Selected = True

                                    'Dim c As New ConfirmingData
                                    'c.PullCustomerINFO(Confirming.Tab, Me.i.Text)

                                    found += 1
                                    foundConfirming += 1
                                    If Confirming.TabControl1.SelectedIndex <> 0 Then
                                        Confirming.TabControl1.SelectedIndex = 0
                                    End If
                                    Confirming.Focus()
                                    'Me.Label5_Click(Nothing, Nothing)
                                    Confirming.Create_Tooltip(Confirming.dpConfirming, "alert")
                                    Exit Sub


                                    'Exit Sub
                                End If

                            Next
                            If Confirming.lvConfirming.SelectedItems.Count = 0 Then
                                Confirming.dpConfirming.Value = was
                            End If
                            was = Confirming.dpSales.Value

                            If foundConfirming = 0 Then

                                Confirming.dpSales.Value = r1.Item(0)
                                For Each Me.i In Confirming.lvSales.Items
                                    If Me.i.Text = Me.lnkID.Text Then
                                        Me.i.Selected = True
                                        found += 1
                                        foundConfirming += 1
                                        If Confirming.TabControl1.SelectedIndex <> 1 Then
                                            Confirming.TabControl1.SelectedIndex = 1
                                        End If
                                        Confirming.Focus()
                                        Confirming.Create_Tooltip(Confirming.dpSales, "alert")
                                        Exit Sub

                                        'Me.Label5_Click(Nothing, Nothing)
                                    Else

                                        'Exit Sub
                                    End If
                                Next
                                If foundConfirming = 0 Then
                                    Confirming.dpSales.Value = was
                                End If
                            End If
                        End While
                        r1.Close()
                        cnn.Close()
                    End If


                    '' run query to see if it is open to form (this case being confirming) if confirming form = true get date, 
                    ''switch date picker then rerun for each and get lead , add tool tip explaining date has been changed to find the lead 
                    '' change below message box to say lead not available for confirming form 

                Case Is = "Sales"
                    Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                    Dim cmdGET As SqlCommand = New SqlCommand("dbo.FindSales", cnn)

                    Dim param1 As SqlParameter = New SqlParameter("@ID", Me.lnkID.Text)
                    cmdGET.Parameters.Add(param1)
                    cmdGET.CommandType = CommandType.StoredProcedure
                    cnn.Open()
                    Dim r1 As SqlDataReader

                    r1 = cmdGET.ExecuteReader
                    Try
                        While r1.Read
                            If r1.Item(0) = True Then
                                If Sales.tbMain.SelectedIndex <> 1 Then
                                    Sales.tbMain.SelectedIndex = 1
                                End If
                                Dim dt As String = r1.Item(1)
                                Dim s = Split(dt, " ")
                                dt = s(0)
                                If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value) <= 0 And DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp2CustomerList.Value) >= 0 Then
                                    If Sales.cboSalesList.SelectedText <> "Unfiltered Sales Dept. List" Then
                                        Sales.cboSalesList.SelectedText = "Unfiltered Sales Dept. List"
                                        For Each Me.i In Sales.lvSales.Items
                                            If Me.i.Text = Me.lnkID.Text Then
                                                Me.i.Selected = True
                                                found += 1
                                                Sales.Focus()
                                                Exit Sub

                                            End If
                                        Next
                                    End If
                                Else
                                    If CType(dt, Date) = Today Then
                                        Sales.cboDateRangeCustomerList.SelectedItem = "Today"
                                    End If
                                    If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value) > 0 Then
                                        Sales.dtp1CustomerList.Value = dt
                                        Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                    End If
                                    If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp2CustomerList.Value) < 0 Then
                                        Sales.dtp2CustomerList.Value = dt
                                        Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                    End If
                                    If Sales.cboSalesList.SelectedItem <> "Unfiltered Sales Dept. List" Then
                                        Sales.cboSalesList.SelectedItem = "Unfiltered Sales Dept. List"
                                    End If
                                    For Each Me.i In Sales.lvSales.Items
                                        If Me.i.Text = Me.lnkID.Text Then
                                            Me.i.Selected = True
                                            found += 1
                                            Sales.Focus()
                                            Exit Sub

                                        End If
                                    Next
                                End If
                                If Sales.tbMain.SelectedIndex <> 1 Then
                                    Sales.tbMain.SelectedIndex = 1
                                End If
                            Else
                                Dim dt As String = r1.Item(2)
                                Dim s = Split(dt, " ")
                                dt = s(0)
                                If r1.Item(1) = True Then
                                    If Sales.cboSalesList.SelectedItem <> "Unconfirmed Appts. For Today" Then
                                        Sales.cboSalesList.SelectedItem = "Unconfirmed Appts. For Today"
                                    End If

                                    'MsgBox(DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value).ToString)
                                    If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value) > 0 Then
                                        Sales.dtp1CustomerList.Value = dt
                                        Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                    End If
                                    If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp2CustomerList.Value) < 0 Then
                                        Sales.dtp2CustomerList.Value = dt
                                        Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                    End If
                                    For Each Me.i In Sales.lvSales.Items
                                        If Me.i.Text = Me.lnkID.Text Then
                                            Me.i.Selected = True
                                            found += 1
                                            Sales.Focus()
                                            Exit Sub
                                        End If
                                    Next
                                End If
                                If Sales.tbMain.SelectedIndex <> 1 Then
                                    Sales.tbMain.SelectedIndex = 1
                                End If
                            End If
                        End While
                    Catch ex As Exception

                    End Try
                Case Is = "WCaller"

                Case Is = "ColdCalling"

                Case Is = "Administration"

                Case Is = "Finance"

                Case Is = "Installation"

                Case Is = "Recovery"

                Case Is = "PreviousCustomer"

                Case Is = "MarketingManager"


            End Select
            If found = 0 And open = True Then
                MsgBox("Cannot find record anywhere in the" & vbCr & "forms your working in!", MsgBoxStyle.Critical, "Alert Link Failed")
            ElseIf open = False Then

                MsgBox("You have no lists open to link to!", MsgBoxStyle.Critical, "Alert Link Failed")
            End If
        End If

        'Dim strLN As String = Me.lnkID.Text
        'Dim strFRM As String = STATIC_VARIABLES.CurrentForm.ToString
        'Select Case strFRM
        '    '' must fill in the rest of the case scenarios
        '    '' ie: administration, wcaller, coldcaller
        '    '' 
        '    Case Is = "CommonNFO"
        '        CommonNFO.RecordID = strLN
        '        CommonNFO.Show()
        '        Exit Select
        '    Case Is = "All"
        '        MsgBox("Place Holder for default 'Admin' startup form.")
        '        Exit Select
        '    Case Is = "Administration"
        '        MsgBox("Place Holder for default 'Administration' form.")

        '        Exit Select
        '    Case Is = "Confirming"
        '        MsgBox("Place Holder for default 'Confirming' form.")
        '        Exit Select
        '    Case Is = "Sales"
        '        MsgBox("Place Holder for default 'Sales Manager' form.")
        '        Exit Select
        'End Select
    End Sub

    Private Sub lnkID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkID.TextChanged
        GetINFO(Me.lnkID.Text)
    End Sub

  



    Private Sub ManageAlerts_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'MsgBox(sender.ToString)
        'Me.WindowState = FormWindowState.Normal
        'Main.Refresh()
        ' Me.Refresh()
    End Sub

 


End Class
