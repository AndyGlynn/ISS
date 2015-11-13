Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class frmCreateList
    Public LoadComplete As Boolean = False
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.cboDateRange.Items.AddRange(New Object() {"All", "Today", "Yesterday", "This Week", "This Week - to date", "This Month", "This Month - to date", "This Year", "This Year - to date", "Next Week", "Next Month", "Last Week", "Last Week - to date", "Last Month", "Last Month - to date", "Last Year", "Last Year - to date", "Custom"})
        Me.cboPCDateRange.Items.AddRange(New Object() {"All", "Today", "Yesterday", "This Week", "This Week - to date", "This Month", "This Month - to date", "This Year", "This Year - to date", "Next Week", "Next Month", "Last Week", "Last Week - to date", "Last Month", "Last Month - to date", "Last Year", "Last Year - to date", "Custom"})
	'' MERGE CODE 10-7-2015
        Dim mpt As New MAPPOINT_LOGIC_V2
        Dim arStates As ArrayList = mpt.Get_Unique_States
        For Each xx As Object In arStates
            Me.cboStateSelection.Items.Add(xx.ToString)
        Next
        '' make sure lbl and cbo are invisible at launch 
        '' 
        Me.lblStateSelect.Visible = False
        Me.cboStateSelection.Visible = False

        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetProducts", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        Try
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            Me.chlstProducts.Items.Clear()

            While r1.Read
                Me.chlstProducts.Items.Add(r1.Item(0))
            End While
            cnn.Close()
        Catch ex As Exception

        End Try
      


    
        For x As Integer = 0 To Me.chlstProducts.Items.Count - 1
            Me.chlstProducts.SetItemChecked(x, True)
        Next
      
        Me.cboDateRange.SelectedItem = Nothing
        Me.cboPCDateRange.SelectedItem = Nothing
        Me.chkWC.Checked = False
        Me.chRehash.Checked = False
        Me.chPC.Checked = False
        Me.chWeekdays.Checked = False
        Me.txtDate1PC.Tag = Me.dpDate1PC.Value.ToShortDateString
        Me.txtDate2PC.Tag = Me.dpDate2PC.Value.ToShortDateString
        Me.txtFrom.Tag = Me.dpFrom.Value.ToShortDateString
        Me.txtTo.Tag = Me.dpTo.Value.ToShortDateString
        Me.txtTimeFrom.Tag = Me.tpFrom.Value.ToShortTimeString
        Me.txtTimeTo.Tag = Me.tpTo.Value.ToShortTimeString
        Me.LoadComplete = True
    End Sub





    Private Sub cboDateRange_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDateRange.SelectedIndexChanged
        If Me.cboDateRange.Text <> "Custom" Then
            If Me.cboDateRange.Text = "All" Then
                Me.dpFrom.Value = "1/1/1900"
                Me.dpTo.Value = "1/1/2100"
                Me.txtFrom.Visible = True
                Me.txtTo.Visible = True
                Exit Sub
            End If

            Me.txtFrom.Visible = False
            Me.txtTo.Visible = False
            Dim d As New DTPManipulation(Me.cboDateRange.Text)
            Me.dpFrom.Text = d.retDateFrom
            Me.dpTo.Text = d.retDateTo

        End If
    End Sub

    Private Sub txtFrom_GotFocus(sender As Object, e As EventArgs) Handles txtFrom.GotFocus

        Me.dpFrom.Focus()

    End Sub

    Private Sub txtTo_GotFocus(sender As Object, e As EventArgs) Handles txtTo.GotFocus

        Me.dpTo.Focus()

    End Sub

    Private Sub dpFrom_GotFocus(sender As Object, e As EventArgs) Handles dpFrom.GotFocus

        If Me.LoadComplete = False Then
            Exit Sub
        End If

        Me.txtFrom.Visible = False
        Me.cboDateRange.SelectedItem = "Custom"
        If Me.dpFrom.Value.ToShortDateString = "1/1/1900" Then
            Me.dpFrom.Value = Today()
        End If

    End Sub

    Private Sub dpTo_GotFocus(sender As Object, e As EventArgs) Handles dpTo.GotFocus
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtTo.Visible = False
        Me.cboDateRange.SelectedItem = "Custom"
        If Me.dpTo.Value.ToShortDateString = "1/1/2100" Then
            Me.dpTo.Value = Today()
        End If

    End Sub

    Private Sub dpTo_ValueChanged(sender As Object, e As EventArgs) Handles dpTo.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtTo.Tag = Me.dpTo.Value.ToShortDateString
    End Sub

    Private Sub dpFrom_ValueChanged(sender As Object, e As EventArgs) Handles dpFrom.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtFrom.Tag = Me.dpFrom.Value.ToShortDateString
    End Sub

    Private Sub dpGenerated_GotFocus(sender As Object, e As EventArgs) Handles dpGenerated.GotFocus
        Me.txtGenerated.Visible = False
        If Me.dpGenerated.Value = "1/1/1900 12:00 AM" Then
            'MsgBox("hit")
            Me.dpGenerated.Value = Today()
        End If


    End Sub

    Private Sub txtGenerated_GotFocus(sender As Object, e As EventArgs) Handles txtGenerated.GotFocus
        Me.dpGenerated.Focus()
    End Sub

    Private Sub dpGenerated_ValueChanged(sender As Object, e As EventArgs) Handles dpGenerated.ValueChanged
        Me.txtGenerated.Tag = Me.dpGenerated.Value.ToShortDateString
    End Sub

    Private Sub txtTimeFrom_GotFocus(sender As Object, e As EventArgs) Handles txtTimeFrom.GotFocus
        Me.tpFrom.Focus()
    End Sub

    Private Sub txtTimeTo_GotFocus(sender As Object, e As EventArgs) Handles txtTimeTo.GotFocus
        Me.tpTo.Focus()
    End Sub

    Private Sub tpFrom_GotFocus(sender As Object, e As EventArgs) Handles tpFrom.GotFocus
        Me.txtTimeFrom.Visible = False
        If Me.tpFrom.Value = "1/1/1900 12:00 AM" Then
            Me.tpFrom.Value = "1/1/1900 10:00 AM"
        End If

    End Sub

    Private Sub tpTo_GotFocus(sender As Object, e As EventArgs) Handles tpTo.GotFocus
        Me.txtTimeTo.Visible = False
        If Me.tpTo.Value = "1/1/1900 11:59 PM" Then
            Me.tpTo.Value = "1/1/1900 8:00 PM"
        End If

    End Sub

   
    Private Sub tpTo_ValueChanged(sender As Object, e As EventArgs) Handles tpTo.ValueChanged
        Me.txtTimeTo.Tag = Me.tpTo.Value.ToShortTimeString
    End Sub

    Private Sub tpFrom_ValueChanged(sender As Object, e As EventArgs) Handles tpFrom.ValueChanged
        Me.txtTimeFrom.Tag = Me.tpFrom.Value.ToShortTimeString
    End Sub

    Private Sub btnCheckProducts_Click(sender As Object, e As EventArgs) Handles btnCheckProducts.Click
        For x As Integer = 0 To Me.chlstProducts.Items.Count - 1
            Me.chlstProducts.SetItemChecked(x, True)
        Next
    End Sub

    Private Sub btnUncheckProducts_Click(sender As Object, e As EventArgs) Handles btnUncheckProducts.Click
        For x As Integer = 0 To Me.chlstProducts.Items.Count - 1
            Me.chlstProducts.SetItemChecked(x, False)
        Next
    End Sub

    Private Sub rdoCity_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCity.CheckedChanged
        Dim rdo As RadioButton = sender
        If rdo.Checked = True Then

            Me.chlstZipCity.Items.Clear()
            Me.txtZipCity.Text = ""
            Me.Label9.Text = "Enter Starting City:"
            Me.Label10.Text = "Show Cities Within"
            Me.Label11.Text = "Miles of Starting City"
            Me.btnShow.Text = "Show Cities"
            Me.lblStateSelect.Visible = True
            Me.cboStateSelection.Visible = True

        End If


    End Sub

    Private Sub rdoZip_CheckedChanged(sender As Object, e As EventArgs) Handles rdoZip.CheckedChanged

        Dim rdo As RadioButton = sender
        If rdo.Checked = True Then
            Me.chlstZipCity.Items.Clear()
            Me.txtZipCity.Text = ""
            Me.Label9.Text = "Enter Starting Zip Code:"
            Me.Label10.Text = "Show Zip Codes Within"
            Me.Label11.Text = "Miles of Starting Zip Code"
            Me.btnShow.Text = "Show Zip Codes"
            Me.lblStateSelect.Visible = False
            Me.cboStateSelection.Visible = False
        End If

    End Sub

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Me.epGeo.Clear()
        If Me.rdoZip.Checked = True And Me.txtZipCity.Text = "" Then
            Me.epGeo.SetError(Me.txtZipCity, "You Must Enter A Starting Zip Code!")
            Exit Sub
        End If
        Dim x As Integer = Me.txtZipCity.TextLength

        If Me.rdoZip.Checked = True And x <> 5 Then
            Me.epGeo.SetError(Me.txtZipCity, "You Must Enter A Valid Zip Code!")
            Exit Sub
        End If
        If Me.rdoZip.Checked = True Then
            For z As Integer = 1 To Me.txtZipCity.TextLength
                Dim c As Char = GetChar(Me.txtZipCity.Text, z)
                If (c = "0") Or (c = "1") Or (c = "2") Or (c = "3") Or (c = "4") Or (c = "5") Or (c = "6") Or (c = "7") Or (c = "8") Or (c = "9") Then
                  Else
                    Me.epGeo.SetError(Me.txtZipCity, "Zip Codes Only Contain Numbers!")
                    Exit Sub
                End If
            Next
        End If
        If Me.rdoCity.Checked = True And Me.txtZipCity.Text = "" Then
            Me.epGeo.SetError(Me.txtZipCity, "You Must Enter A Starting City!")
            Exit Sub
        End If
        If Me.numMiles.Value = 0 Then
            Me.epGeo.SetError(Me.numMiles, "Radius Search Must Be 1 Mile Or More!")
            Exit Sub
        End If
        ''Call Mappoint search for city, if doesnt exist fire code below
        ''msgbox("Cannot Find City!)
        ''Me.txtZipCity.SelectAll()
        ''If more than 1 exists fire form with listview containing city and state to select appropriate city then click ok button and continue mappoint query


	'' MERGE CODE: 10-7-2015
	'' 
        ''    Add Mappoint query code to populate chlstZipCity then loop to check all

        Dim mpt As New MAPPOINT_LOGIC_V2
        Select Case btnShow.Text
            Case "Show Zip Codes"
                Me.chlstZipCity.Items.Clear()
                Dim arResults As New ArrayList
                Dim BeginList As List(Of MAPPOINT_LOGIC_V2.ZipState)
                BeginList = mpt.GetUniqueZipState
                Me.Cursor = Cursors.WaitCursor
                Dim StartState As String = mpt.Get_STATE_ForZip(Me.txtZipCity.Text)
                arResults = mpt.Search_Radius_ZipCode(BeginList, Me.numMiles.Value, Me.txtZipCity.Text, StartState)
                Dim g As Object
                For Each g In arResults
                    Me.chlstZipCity.Items.Add(g.ToString)
                Next
                Me.Cursor = Cursors.Default
                Exit Select
            Case "Show Cities"
                Me.chlstZipCity.Items.Clear()
                Dim arResults As New ArrayList
                Dim BeginList As List(Of MAPPOINT_LOGIC_V2.CityState)
                BeginList = mpt.GetUniqueCityState
                Me.Cursor = Cursors.WaitCursor
                Dim Start_State As String = mpt.Get_STATE_ForCity(Me.txtZipCity.Text)
                arResults = mpt.Search_Radius_City(BeginList, Me.numMiles.Value, Me.txtZipCity.Text, Me.cboStateSelection.Text)
                Dim g As Object
                For Each g In arResults
                    Me.chlstZipCity.Items.Add(g.ToString)
                Next
                Me.Cursor = Cursors.Default
                Exit Select
        End Select


    End Sub

    Private Sub btnCheckZip_Click(sender As Object, e As EventArgs) Handles btnCheckZip.Click
        For x As Integer = 0 To Me.chlstZipCity.Items.Count - 1
            Me.chlstZipCity.SetItemChecked(x, True)
        Next
    End Sub

    Private Sub btnUncheckZip_Click(sender As Object, e As EventArgs) Handles btnUncheckZip.Click
        For x As Integer = 0 To Me.chlstZipCity.Items.Count - 1
            Me.chlstZipCity.SetItemChecked(x, False)
        Next
    End Sub

    Private Sub chkWC_CheckedChanged(sender As Object, e As EventArgs) Handles chkWC.CheckedChanged
        If Me.chkWC.Checked = True Then
            Me.chlstWC.Enabled = True
            For x As Integer = 0 To Me.chlstWC.Items.Count - 1
                Me.chlstWC.SetItemChecked(x, True)
            Next
        Else
            Me.chlstWC.Enabled = False
            For x As Integer = 0 To Me.chlstWC.Items.Count - 1
                Me.chlstWC.SetItemChecked(x, False)
            Next
        End If
    End Sub

    Private Sub btnCheckWC_Click(sender As Object, e As EventArgs) Handles btnCheckWC.Click
        If Me.chlstWC.Enabled = False Then
            Exit Sub
        End If
        For x As Integer = 0 To Me.chlstWC.Items.Count - 1
            Me.chlstWC.SetItemChecked(x, True)
        Next
    End Sub

    Private Sub btnUncheckWC_Click(sender As Object, e As EventArgs) Handles btnUncheckWC.Click
        If Me.chlstWC.Enabled = False Then
            Exit Sub
        End If
        For x As Integer = 0 To Me.chlstWC.Items.Count - 1
            Me.chlstWC.SetItemChecked(x, False)
        Next
    End Sub

    Private Sub chPC_CheckedChanged(sender As Object, e As EventArgs) Handles chPC.CheckedChanged
        If Me.chPC.Checked = True Then
            Me.GroupBox6.Enabled = True
            Me.chFutureInterest.Enabled = True
            Me.chLoanSatisfied.Enabled = True
            Me.chApprovedFor.Enabled = True
            Me.txtApprovedDollars.Enabled = True
            Me.numMonths.Enabled = True
            Me.cboPCDateRange.Enabled = True
            Me.dpDate1PC.Enabled = True
            Me.dpDate2PC.Enabled = True
            Me.txtDate1PC.Enabled = True
            Me.txtDate2PC.Enabled = True
            Me.Label16.Enabled = True
            Me.Label17.Enabled = True
            Me.Label18.Enabled = True
            Me.Label19.Enabled = True
            Me.numReferences.Enabled = True
            Me.rdoAll.Checked = True
        
        Else
            Me.GroupBox6.Enabled = False
            Me.rdoAll.Checked = True
            Me.chFutureInterest.Enabled = False
            Me.chFutureInterest.Checked = False
            Me.chLoanSatisfied.Enabled = False
            Me.chLoanSatisfied.Checked = False
            Me.chApprovedFor.Enabled = False
            Me.chApprovedFor.Checked = False
            Me.txtApprovedDollars.Enabled = False
            Me.txtApprovedDollars.Text = ""
            Me.numMonths.Enabled = False
            Me.numMonths.Value = 0
            Me.cboPCDateRange.Enabled = False
            Me.cboPCDateRange.SelectedItem = Nothing
            Me.dpDate1PC.Enabled = False
            Me.dpDate1PC.Value = "1/1/1900 12:00:00"
            Me.dpDate2PC.Enabled = False
            Me.dpDate2PC.Value = "1/1/2100 12:00:00"
            Me.txtDate1PC.Enabled = False
            Me.txtDate1PC.Visible = True
            Me.txtDate2PC.Enabled = False
            Me.txtDate2PC.Visible = True
            Me.Label16.Enabled = False
            Me.Label17.Enabled = False
            Me.Label18.Enabled = False
            Me.Label19.Enabled = False
            Me.numReferences.Value = 0.0
            Me.numReferences.Enabled = False
       
        End If
    End Sub

    Private Sub chRehash_CheckedChanged(sender As Object, e As EventArgs) Handles chRehash.CheckedChanged
        If Me.chRehash.Checked = True Then
            Me.chlstRehash.Enabled = True
            Me.Label26.Enabled = True
            Me.Label27.Enabled = True
            Me.numPar.Enabled = True
            Me.txtQuoted.Enabled = True
            For x As Integer = 0 To Me.chlstRehash.Items.Count - 1
                Me.chlstRehash.SetItemChecked(x, True)
            Next
        Else
            Me.chlstRehash.Enabled = False
            Me.Label26.Enabled = False
            Me.Label27.Enabled = False
            Me.numPar.Enabled = False
            Me.txtQuoted.Enabled = False
            Me.txtQuoted.Text = ""
            Me.numPar.Value = 0
            For x As Integer = 0 To Me.chlstRehash.Items.Count - 1
                Me.chlstRehash.SetItemChecked(x, False)
            Next
        End If
    End Sub

    Private Sub btnCheckRehash_Click(sender As Object, e As EventArgs) Handles btnCheckRehash.Click
        If Me.chRehash.Checked = True Then
            For x As Integer = 0 To Me.chlstRehash.Items.Count - 1
                Me.chlstRehash.SetItemChecked(x, True)
            Next
        End If
    End Sub

    Private Sub btnUncheckRehash_Click(sender As Object, e As EventArgs) Handles btnUncheckRehash.Click
        If Me.chRehash.Checked = True Then
            For x As Integer = 0 To Me.chlstRehash.Items.Count - 1
                Me.chlstRehash.SetItemChecked(x, False)
            Next
        End If
    End Sub

    Private Sub cboPCDateRange_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPCDateRange.SelectedIndexChanged
        If Me.cboPCDateRange.Text <> "Custom" Then
            If Me.cboPCDateRange.Text = "All" Then
                Me.dpDate1PC.Value = "1/1/1900"
                Me.dpDate2PC.Value = "1/1/2100"
                Me.txtDate1PC.Visible = True
                Me.txtDate2PC.Visible = True
                Exit Sub
            End If
            Me.txtDate1PC.Visible = False
            Me.txtDate2PC.Visible = False
            Dim d As New DTPManipulation(Me.cboPCDateRange.Text)
            Me.dpDate1PC.Text = d.retDateFrom
            Me.dpDate2PC.Text = d.retDateTo
        End If
    End Sub

    Private Sub dpDate1PC_GotFocus(sender As Object, e As EventArgs) Handles dpDate1PC.GotFocus
        If Me.LoadComplete = False Then
            Exit Sub
        End If

        Me.txtDate1PC.Visible = False
        Me.cboPCDateRange.SelectedItem = "Custom"
        Me.dpDate1PC.Value = Today()
    End Sub

    Private Sub dpDate2PC_GotFocus(sender As Object, e As EventArgs) Handles dpDate2PC.GotFocus
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtDate2PC.Visible = False
        Me.cboPCDateRange.SelectedItem = "Custom"
        Me.dpDate2PC.Value = Today()
    End Sub

    Private Sub txtDate1PC_GotFocus(sender As Object, e As EventArgs) Handles txtDate1PC.GotFocus
        Me.dpDate1PC.Focus()
    End Sub

    Private Sub txtDate2PC_GotFocus(sender As Object, e As EventArgs) Handles txtDate2PC.GotFocus
        Me.dpDate2PC.Focus()
    End Sub

    Private Sub dpDate1PC_ValueChanged(sender As Object, e As EventArgs) Handles dpDate1PC.ValueChanged
        Me.txtDate1PC.Tag = Me.dpDate1PC.Value.ToShortDateString
    End Sub

    Private Sub dpDate2PC_ValueChanged(sender As Object, e As EventArgs) Handles dpDate2PC.ValueChanged
        Me.txtDate2PC.Tag = Me.dpDate2PC.Value.ToShortDateString
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Me.RichTextBox1.Text = "" ''Remove After Testing 


        Me.tabDate.ImageKey = Nothing
        Me.tabDate.ToolTipText = ""
        Me.tabProducts.ImageKey = Nothing
        Me.tabProducts.ToolTipText = ""
        Me.tabGeo.ImageKey = Nothing
        Me.tabGeo.ToolTipText = ""
        Me.tabWC.ImageKey = Nothing
        Me.tabWC.ToolTipText = ""
        Me.tabPC.ImageKey = Nothing
        Me.tabPC.ToolTipText = ""
        Me.tabRecovery.ImageKey = Nothing
        Me.tabRecovery.ToolTipText = ""
        Me.tabMarketers.ImageKey = Nothing
        Me.tabMarketers.ToolTipText = ""

        Me.epForm.Clear()
        Dim tabDate As Integer = 0
        Dim tabProd As Integer = 0
        Dim tabGeo As Integer = 0
        Dim tabWC As Integer = 0
        Dim tabPC As Integer = 0
        Dim tabRehash As Integer = 0
        Dim tabMarketer As Integer = 0

        ''Date/Time Tab 
        If Me.dpFrom.Value.ToShortDateString <> "1/1/1900" And Me.dpTo.Value.ToShortDateString = "1/1/2100" Then              '' Change If Statement to reflect datepicker value instead of visible properties
            Me.epForm.SetError(Me.dpTo, "You Must Supply a Date!")
            tabDate += 1
        End If
        If Me.dpFrom.Value.ToShortDateString = "1/1/1900" And Me.dpTo.Value.ToShortDateString <> "1/1/2100" Then              '' Change If Statement to reflect datepicker value instead of visible properties
            Me.epForm.SetError(Me.dpFrom, "You Must Supply a Date!")
            tabDate += 1
        End If
        If Me.tpFrom.Value.ToShortTimeString = "12:00 AM" And Me.tpTo.Value.ToShortTimeString <> "11:59 PM" Then
            Me.epForm.SetError(Me.tpFrom, "You Must Supply a Time!")
            tabDate += 1
        End If
        If Me.tpFrom.Value.ToShortTimeString <> "12:00 AM" And Me.tpTo.Value.ToShortTimeString = "11:59 PM" Then
            Me.epForm.SetError(Me.tpTo, "You Must Supply a Time!")
            tabDate += 1
        End If

        ''Marketers Tab 
        If Me.chlstMarketers.CheckedItems.Count <= 0 Then
            Me.epForm.SetError(Me.chlstMarketers, "You Must Select at Least One Marketer!")
            tabMarketer += 1
        End If
        ''Products Tab
        If Me.chlstProducts.CheckedItems.Count <= 0 Then
            Me.epForm.SetError(Me.chlstProducts, "You Must Select at Least One Product!")
            tabProd += 1
        End If

        ''Geo Tab 
        If Me.chlstZipCity.Items.Count > 0 And Me.chlstZipCity.CheckedItems.Count <= 0 Then
            If Me.rdoZip.Checked = True Then
                Me.epForm.SetError(Me.chlstProducts, "You Must Select at Least One Zip Code!")
                tabProd += 1
            Else
                Me.epForm.SetError(Me.chlstProducts, "You Must Select at Least One City!")

                Me.epForm.SetIconPadding(Me.chlstProducts, 20)
                tabProd += 1
            End If
        End If

        ''Warm Calling Tab
        If Me.chlstWC.CheckedItems.Count <= 0 And Me.chkWC.Checked = True Then
            Me.epForm.SetError(Me.chlstWC, "You Must Select at Least One Marketing Result!")
            tabWC += 1
        End If

        ''Previous Customer Tab
        If Me.chPC.Checked = True Then
            If Me.chApprovedFor.Checked = True And Me.txtApprovedDollars.Text = "" Then
                Me.epForm.SetError(Me.txtApprovedDollars, "You Must Supply a Dollar Amount! ")
                tabPC += 1
            End If
            If Me.chApprovedFor.Checked = True Then
                For z As Integer = 1 To Me.txtApprovedDollars.TextLength
                    Dim c As Char = GetChar(Me.txtApprovedDollars.Text, z)
                    If (c = "0") Or (c = "1") Or (c = "2") Or (c = "3") Or (c = "4") Or (c = "5") Or (c = "6") Or (c = "7") Or (c = "8") Or (c = "9") Then
                    Else
                        Me.epGeo.SetError(Me.txtApprovedDollars, "Field Can Only Contain Numbers!")
                        tabPC += 1

                    End If
                Next
            End If
            If Me.dpDate1PC.Value.ToShortDateString = "1/1/1900" And Me.dpDate2PC.Value.ToShortDateString <> "1/1/2100" Then
                Me.epForm.SetError(Me.dpDate1PC, "You Must Supply a Date!")
                tabPC += 1
            End If
            If Me.dpDate1PC.Value.ToShortDateString <> "1/1/1900" And Me.dpDate2PC.Value.ToShortDateString = "1/1/2100" Then
                Me.epForm.SetError(Me.dpDate2PC, "You Must Supply a Date!")
                tabPC += 1
            End If
        End If

        ''Rehash Tab
        If Me.chRehash.Checked = True Then
            If Me.chlstRehash.CheckedItems.Count <= 0 Then
                Me.epForm.SetError(Me.chlstRehash, "You Must Select at Least One Result!")
                tabRehash += 1
            End If
            For z As Integer = 1 To Me.txtQuoted.TextLength
                Dim c As Char = GetChar(Me.txtQuoted.Text, z)
                If (c = "0") Or (c = "1") Or (c = "2") Or (c = "3") Or (c = "4") Or (c = "5") Or (c = "6") Or (c = "7") Or (c = "8") Or (c = "9") Then
                Else
                    Me.epGeo.SetError(Me.txtQuoted, "Field Can Only Contain Numbers!")
                    tabRehash += 1
                End If
            Next
        End If
        Dim tab As Integer = 0
        If tabDate > 0 Then
            Me.tabDate.ImageIndex = 0
            Me.tabDate.ToolTipText = "Errors on this Tab"
            tab += 1
        End If
        If tabProd > 0 Then
            Me.tabProducts.ImageIndex = 0
            Me.tabProducts.ToolTipText = "Errors on this Tab"
            tab += 1
        End If
        If tabGeo > 0 Then
            Me.tabGeo.ImageIndex = 0
            Me.tabGeo.ToolTipText = "Errors on this Tab"
            tab += 1
        End If
        If tabWC > 0 Then
            Me.tabWC.ImageIndex = 0
            Me.tabWC.ToolTipText = "Errors on this Tab"
            tab += 1
        End If
        If tabPC > 0 Then
            Me.tabPC.ImageIndex = 0
            Me.tabPC.ToolTipText = "Errors on this Tab"
            tab += 1
        End If
        If tabRehash > 0 Then
            Me.tabRecovery.ImageIndex = 0
            Me.tabRecovery.ToolTipText = "Errors on this Tab"
            tab += 1
        End If
        If tabMarketer > 0 Then
            Me.tabMarketers.ImageIndex = 0
            Me.tabMarketers.ToolTipText = "Errors on this Tab"
            tab += 1
        End If
        If tab > 0 Then
            Exit Sub
        End If

        ''ADD SQL QUERY CODE HERE 
        'MsgBox("Creating List")

        ''Date Tab
        Dim Where As String = ""
        Dim DateQuery As String = ""
        Dim MarkProdGeo As String = ""
        Dim WCQuery As String = ""
        Dim Rehash As String = ""
        Dim PCQuery As String = ""
        Dim Slct As String = "Select Distinct(Enterlead.ID), Contact1LastName, Contact1FirstName, Contact2LastName, Contact2FirstName, StAddress, City, State, Zip, HousePhone, AltPhone1, AltPhone2, Enterlead.Product1, Enterlead.Product2, Enterlead.Product3, LeadGeneratedOn, ApptDate, ApptDay, ApptTime, Result, MarketingResults, IsPreviousCustomer, IsRecovery "
        Dim Orderby As String = ""
        Dim checked As Integer = 0
  
        ''Add columns depending on orderby clause 
        For x As Integer = 0 To Me.chlstOrderBy.CheckedItems.Count - 1
            Select Case Me.chlstOrderBy.CheckedItems(x).ToString
            Case Is = "Reference Rating"
                    Slct = Slct & ", ReferenceRating"
                    checked += 1
                Case Is = "Sale Closed Date"
                    Slct = Slct & ", JobClosed"
                    checked += 1
                Case Is = "Loan Satisfied Date"
                    Slct = Slct & ", ExpectedPayOff"
                Case Is = "Approved Loan Amount"
                    Slct = Slct & ", ApprovedFor"
            End Select
        Next

        Slct = Slct & " From Enterlead"


        If Me.chPC.Checked = True Then
            If Me.chLoanSatisfied.Checked = True Or Me.chApprovedFor.Checked = True Then
                Slct = Slct & " Join tblFinance on Enterlead.ID = tblFinance.LeadNum"
            End If
            If Me.numReferences.Value > 0 Or Me.dpDate1PC.Value.ToShortDateString <> "1/1/1900" Or checked > 0 Then
                Slct = Slct & " Join SaleDetail on Enterlead.ID = SaleDetail.LeadNum"
            End If
        End If


        Dim Fby As Integer = 0
        If Me.dpGenerated.Value.ToShortDateString <> "1/1/1900" Then
            DateQuery = "LeadGeneratedOn >= '" & Me.dpGenerated.Value.ToString & "'"
            Fby += 1
        End If
        If Me.dpFrom.Value.ToShortDateString <> "1/1/1900" Then
            If DateQuery <> "" Then
                DateQuery = DateQuery & " and "
            End If
            DateQuery = DateQuery & "ApptDate Between '" & Me.dpFrom.Value.ToString & "' and '" & Me.dpTo.Value.ToString & "'"
            Fby += 1
        End If
        If Me.tpFrom.Value.ToShortTimeString <> "12:00 AM" Then
            If DateQuery <> "" Then
                DateQuery = DateQuery & " and "
            End If
            DateQuery = DateQuery & "ApptTime Between '" & Me.tpFrom.Value.ToString & "' And '" & Me.tpTo.Value.ToString & "'"
            Fby += 1
        End If
        If Me.chWeekdays.Checked = True Then
            If DateQuery <> "" Then
                DateQuery = DateQuery & " and "
            End If
            DateQuery = DateQuery & "ApptDay <> 'Saturday' and ApptDay <> 'Sunday'"
            Fby += 1
        End If
        ''Marketer Tab 
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand
        Dim r1 As SqlDataReader
        cmdGet = New SqlCommand("Select Count(distinct(marketer)) from EnterLead", cnn)


        cmdGet.CommandType = CommandType.Text
        cnn.Open()
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r1.Read()
        If r1.Item(0) > Me.chlstMarketers.Items.Count Then    '' If true were working with a filtered down set of Marketers 
            If MarkProdGeo <> "" Then
                MarkProdGeo = MarkProdGeo & " and "
            End If
            For x As Integer = 0 To Me.chlstMarketers.CheckedItems.Count - 1
                If x = 0 Then
                    MarkProdGeo = MarkProdGeo & "("
                End If
                If x <> Me.chlstMarketers.CheckedItems.Count - 1 Then

                    MarkProdGeo = MarkProdGeo & "Marketer = '" & Me.chlstMarketers.CheckedItems(x).ToString & "' and "    ''Add Comma after each item until last item 
                    Fby += 1
                Else
                    MarkProdGeo = MarkProdGeo & "Marketer = '" & Me.chlstMarketers.CheckedItems(x).ToString & "')"       ''Last Item Dont Add comma to end of where 
                    Fby += 1
                End If
            Next
        Else
            If Me.chlstMarketers.Items.Count <> Me.chlstMarketers.CheckedItems.Count Then  ''If all items are checked in list checkbox and were not working with a filtered down marketer list then no need to query by marketer at all
                If MarkProdGeo <> "" Then
                    MarkProdGeo = MarkProdGeo & " and "
                End If
                For x As Integer = 0 To Me.chlstMarketers.CheckedItems.Count - 1
                    If x = 0 Then
                        MarkProdGeo = MarkProdGeo & "("
                    End If
                    If x <> Me.chlstMarketers.CheckedItems.Count - 1 Then

                        MarkProdGeo = MarkProdGeo & "Marketer = '" & Me.chlstMarketers.CheckedItems(x).ToString & "' or "    ''Add Comma after each item until last item 
                        Fby += 1
                    Else
                        MarkProdGeo = MarkProdGeo & "Marketer = '" & Me.chlstMarketers.CheckedItems(x).ToString & "')"       ''Last Item Dont Add comma to end of where 
                        Fby += 1
                    End If
                Next
            End If
        End If

        ''Products
        If Me.chlstProducts.Items.Count <> Me.chlstProducts.CheckedItems.Count Then ''If true filter by checked products, if false no need to filter products (skip to next step)
            If MarkProdGeo <> "" Then
                MarkProdGeo = MarkProdGeo & " and "
            End If
            For x As Integer = 0 To Me.chlstProducts.CheckedItems.Count - 1
                If x = 0 Then
                    MarkProdGeo = MarkProdGeo & "("
                End If
                If x <> Me.chlstProducts.CheckedItems.Count - 1 Then

                    MarkProdGeo = MarkProdGeo & "(" & "Enterlead.Product1 = '" & Me.chlstProducts.CheckedItems(x).ToString & "' or Enterlead.Product2 = '" & Me.chlstProducts.CheckedItems(x).ToString & "' or Enterlead.Product3 = '" & Me.chlstProducts.CheckedItems(x).ToString & "') or "    ''Add Comma after each item until last item 
                    Fby += 1
                Else
                    MarkProdGeo = MarkProdGeo & "(" & "Enterlead.Product1 = '" & Me.chlstProducts.CheckedItems(x).ToString & "' or Enterlead.Product2 = '" & Me.chlstProducts.CheckedItems(x).ToString & "' or Enterlead.Product3 = '" & Me.chlstProducts.CheckedItems(x).ToString & "'))"       ''Last Item Dont Add comma to end of where 
                    Fby += 1
                End If
            Next
        End If

        ''Geo Tab 
        If Me.chlstZipCity.CheckedItems.Count > 0 Then
            If MarkProdGeo <> "" Then
                MarkProdGeo = MarkProdGeo & " and "
            End If
            Dim CorZ As String = ""
            If Me.rdoZip.Checked = True Then
                CorZ = "Zip"
            Else
                CorZ = "City"
            End If
            For x As Integer = 0 To Me.chlstZipCity.CheckedItems.Count - 1
                If x = 0 Then
                    MarkProdGeo = MarkProdGeo & "("
                End If
                If x <> Me.chlstZipCity.CheckedItems.Count - 1 Then

                    MarkProdGeo = MarkProdGeo & CorZ & " = '" & Me.chlstZipCity.CheckedItems(x).ToString & "' or "    ''Add Comma after each item until last item 
                    Fby += 1
                Else
                    MarkProdGeo = MarkProdGeo & CorZ & " = '" & Me.chlstZipCity.CheckedItems(x).ToString & "')"       ''Last Item Dont Add comma to end of where 
                    Fby += 1
                End If
            Next
        End If

        '' Warm Calling Tab 

        If Me.chkWC.Checked = True Then
            If WCQuery <> "" Then
                WCQuery = WCQuery & " and "
            End If
            For x As Integer = 0 To Me.chlstWC.CheckedItems.Count - 1
                If x = 0 Then
                    WCQuery = WCQuery & "("
                End If
                Dim Field As String = ""
                If Me.chlstWC.CheckedItems(x) = "Reset" Or Me.chlstWC.CheckedItems(x) = "Not Hit" Or Me.chlstWC.CheckedItems(x) = "Not Issued" Then
                    Field = "Result"
                Else
                    Field = "MarketingResults"
                End If
                If x <> Me.chlstWC.CheckedItems.Count - 1 Then

                    WCQuery = WCQuery & "(" & Field & " = '" & Me.chlstWC.CheckedItems(x).ToString & "' and isRecovery = 'False' and isPreviousCustomer = 'False') or "    ''Add Comma after each item until last item 
                    Fby += 1
                Else
                    WCQuery = WCQuery & "(" & Field & " = '" & Me.chlstWC.CheckedItems(x).ToString & "' and isRecovery = 'False' and isPreviousCustomer = 'False'))"       ''Last Item Dont Add comma to end of where 
                    Fby += 1
                End If
            Next
        End If

        '' Recovery
        If Me.chRehash.Checked = True Then
            Dim Per As Double = 0
            If Me.numPar.Value = 0 Then
                Per = 100

            End If
            Per = Me.numPar.Value / 100
            Per = 1 - Per
            Dim Q As Integer = 0
            If Me.txtQuoted.Text <> "" Then
                Q = CType(Me.txtQuoted.Text, Integer)
            End If

            For x As Integer = 0 To Me.chlstRehash.CheckedItems.Count - 1
                If x = 0 Then
                    Rehash = Rehash & "("
                End If
                Dim Field As String = ""
                If Me.chlstRehash.CheckedItems(x) = "Reset" Or Me.chlstRehash.CheckedItems(x) = "Not Hit" Or Me.chlstRehash.CheckedItems(x) = "Not Issued" Or Me.chlstRehash.CheckedItems(x) = "Demo/No Sale" Or Me.chlstRehash.CheckedItems(x) = "Recission Cancel" Then
                    Field = "Result"
                Else
                    Field = "MarketingResults"
                End If
                If x <> Me.chlstRehash.CheckedItems.Count - 1 Then

                    Rehash = Rehash & "(" & Field & " = '" & Me.chlstRehash.CheckedItems(x).ToString & "' and isRecovery = 'True' and isPreviousCustomer = 'False' and (QuotedSold *" & Per & ") >= ParPrice and QuotedSold >= " & Q & " ) or "    ''Add Comma after each item until last item 
                    Fby += 1
                Else
                    Rehash = Rehash & "(" & Field & " = '" & Me.chlstRehash.CheckedItems(x).ToString & "' and isRecovery = 'True' and isPreviousCustomer = 'False' and (QuotedSold *" & Per & ") >= ParPrice and QuotedSold >= " & Q & "))"       ''Last Item Dont Add comma to end of where 
                    Fby += 1
                End If
            Next




        End If

        ''Reconstruct Where Statement 
        If WCQuery <> "" Then ''Constructs Query For Warm Calling 
            If DateQuery <> "" And MarkProdGeo <> "" Then
                Where = "(" & DateQuery & " and " & MarkProdGeo & " and " & WCQuery & ")"
            ElseIf DateQuery <> "" And MarkProdGeo = "" Then
                Where = "(" & DateQuery & " and " & WCQuery & ")"
            ElseIf DateQuery = "" And MarkProdGeo <> "" Then
                Where = "(" & MarkProdGeo & " and " & WCQuery & ")"
            ElseIf DateQuery = "" And MarkProdGeo = "" Then
                Where = WCQuery
            End If
        End If

        If Where <> "" And Rehash <> "" Then
            Where = Where & " or "
        End If

        If Rehash <> "" Then ''Constructs Query For Rehash 
            If DateQuery <> "" And MarkProdGeo <> "" Then
                Where = Where & "(" & DateQuery & " and " & MarkProdGeo & " and " & Rehash & ")"
            ElseIf DateQuery <> "" And MarkProdGeo = "" Then
                Where = Where & "(" & DateQuery & " and " & Rehash & ")"
            ElseIf DateQuery = "" And MarkProdGeo <> "" Then
                Where = Where & "(" & MarkProdGeo & " and " & Rehash & ")"
            ElseIf DateQuery = "" And MarkProdGeo = "" Then
                Where = Where & Rehash
            End If
        End If

        If Me.chPC.Checked = True Then ''Build PC Query Here 

            If Me.chFutureInterest.Checked = True Then
                PCQuery = PCQuery & "(Enterlead.Product1 <> '' and Enterlead.Product2 <> '' and Enterlead.Product3 <> '')"
            End If

            If Me.rdoCash.Checked = True Then
                If PCQuery <> "" Then
                    PCQuery = PCQuery & " and "
                End If
                PCQuery = PCQuery & "Enterlead.Cash = 'True'"
            End If

            If Me.rdoLoan.Checked = True Then
                If PCQuery <> "" Then
                    PCQuery = PCQuery & " and "
                End If
                PCQuery = PCQuery & "Enterlead.Finance = 'True'"
            End If

            If Me.chApprovedFor.Checked = True Then
                If PCQuery <> "" Then
                    PCQuery = PCQuery & " and "
                End If
                PCQuery = PCQuery & "ApprovedFor >= " & CType(Me.txtApprovedDollars.Text, Integer)
            End If

            If Me.chLoanSatisfied.Checked = True Then
                If PCQuery <> "" Then
                    PCQuery = PCQuery & " and "
                End If
                Dim Dt As Date = Today
                Dt = Dt.Month & "/1/" & Dt.Year
                If Me.numMonths.Value > 0 Then
                    Dt = Dt.AddMonths(Me.numMonths.Value)
                End If
                PCQuery = PCQuery & "ExpectedPayOff <= '" & Dt.ToShortDateString & "'"
            End If

            If Me.numReferences.Value > 0 Then
                If PCQuery <> "" Then
                    PCQuery = PCQuery & " and "
                End If
                PCQuery = PCQuery & "ReferenceRating >= " & Me.numReferences.Value
            End If

            If Me.dpDate1PC.Value.ToShortDateString <> "1/1/1900" Then
                If PCQuery <> "" Then
                    PCQuery = PCQuery & " and "
                End If
                PCQuery = PCQuery & "JobClosed Between '" & Me.dpDate1PC.Value.ToShortDateString & "' and '" & Me.dpDate2PC.Value.ToShortDateString & "'"
            End If

            If PCQuery <> "" Then
                PCQuery = PCQuery & " and isPreviousCustomer = 'True'"
            End If

        End If




        If PCQuery <> "" And (WCQuery <> "" Or Rehash <> "") Then
            Where = Where & " or "
        End If

        If Me.chPC.Checked = True Then


            If PCQuery <> "" Then ''Constructs Query For PC's 
                If MarkProdGeo <> "" Then
                    Where = Where & "(" & MarkProdGeo & " and " & PCQuery & ")"
                Else
                    Where = Where & PCQuery
                End If
            Else
                If MarkProdGeo <> "" And Me.chPC.Checked = True Then
                    Where = Where & "(" & MarkProdGeo & " and " & "isPreviousCustomer = 'True')"
                Else
                    If Where = "" Then
                        Where = "isPreviousCustomer = 'True'"
                    Else
                        If PCQuery = "" Then
                            Where = Where & " or isPreviousCustomer = 'True'"
                        End If
                    End If
                End If
            End If
        End If

        If Where <> "" Then
            Where = " Where " & Where
        End If
        If Me.chkWC.Checked = False And Me.chRehash.Checked = False And Me.chPC.Checked = False Then
            MsgBox("You Must Check At Least One Type Of Lead To Generate List!", MsgBoxStyle.Exclamation, "No Lead Types Checked")
            Exit Sub
        End If
        Slct = Slct & Where
        ''Add Order by Here
        If Me.chGroupBy.Checked = True Then
            Orderby = "ispreviouscustomer asc, isrecovery asc, "
        End If
        For x As Integer = 0 To Me.chlstOrderBy.CheckedItems.Count - 1
            Select Case Me.chlstOrderBy.CheckedItems(x).ToString
                Case Is = "City, State"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "City asc"
                    Else
                        Orderby = Orderby & "City asc, "
                    End If
                Case Is = "Zip Code"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "Zip desc"
                    Else
                        Orderby = Orderby & "Zip desc, "
                    End If
                Case Is = "Generated On"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "LeadGeneratedOn desc"
                    Else
                        Orderby = Orderby & "LeadGeneratedOn desc, "
                    End If
                Case Is = "Appointment Date"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "ApptDate desc"
                    Else
                        Orderby = Orderby & "ApptDate desc, "
                    End If
                Case Is = "Appointment Time"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "ApptTime desc"
                    Else
                        Orderby = Orderby & "ApptTime desc, "
                    End If
                Case Is = "Products"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "Enterlead.Product1 asc"
                    Else
                        Orderby = Orderby & "EnterLead.Product1 asc, "
                    End If
                Case Is = "Reference Rating"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "ReferenceRating desc"
                    Else
                        Orderby = Orderby & "ReferenceRating desc, "
                    End If
                Case Is = "Cash"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "Enterlead.Cash desc"
                    Else
                        Orderby = Orderby & "EnterLead.Cash desc, "
                    End If
                Case Is = "Finance"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "Enterlead.Finance desc"
                    Else
                        Orderby = Orderby & "EnterLead.Finance desc, "
                    End If
                Case Is = "Sale Closed Date"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "JobClosed desc"
                    Else
                        Orderby = Orderby & "JobClosed desc, "
                    End If
                Case Is = "Loan Satisfied Date"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "ExpectedPayOff desc"
                    Else
                        Orderby = Orderby & "ExpectedPayOff desc, "
                    End If
                Case Is = "Approved Loan Amount"
                    If x = Me.chlstOrderBy.CheckedItems.Count - 1 Then
                        Orderby = Orderby & "ApprovedFor desc"
                    Else
                        Orderby = Orderby & "ApprovedFor desc, "
                    End If
            End Select
        Next
        If Orderby <> "" Then
            Orderby = " Order By " & Orderby
            Slct = Slct & Orderby
        End If


        Me.RichTextBox1.Text = Slct

        '''Done Now Must Debug the shit out of Select Statement 
        '''Also need to Add Order By Statements to end 
        ''' 

        Dim cl As New createListPrintOperations
        cl.CreateWireFrameHTML(Slct)

    End Sub

    Private Sub chLoanSatisfied_CheckedChanged(sender As Object, e As EventArgs) Handles chLoanSatisfied.CheckedChanged
        If Me.chLoanSatisfied.Checked = False Then
            Me.numMonths.Value = 0
        End If
    End Sub

    Private Sub chApprovedFor_CheckedChanged(sender As Object, e As EventArgs) Handles chApprovedFor.CheckedChanged
        If Me.chApprovedFor.Checked = False Then
            Me.txtApprovedDollars.Text = ""
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Me.epForm.Clear()
        Me.epGeo.Clear()
        Me.cboDateRange.SelectedItem = Nothing
        Me.cboPCDateRange.SelectedItem = Nothing
        Me.dpGenerated.Value = "1/1/1900 12:00 AM"
        Me.txtGenerated.Visible = True
        Me.dpFrom.Value = "1/1/1900 12:00 AM"
        Me.txtFrom.Visible = True
        Me.dpTo.Value = "1/1/2100 12:00 AM"
        Me.txtTo.Visible = True
        Me.tpFrom.Value = "1/1/1900 12:00 AM"
        Me.txtTimeFrom.Visible = True
        Me.tpTo.Value = "1/1/1900 11:59 PM"
        Me.txtTimeTo.Visible = True
        Me.chWeekdays.Checked = False
        For x As Integer = 0 To Me.chlstProducts.Items.Count - 1
            Me.chlstProducts.SetItemChecked(x, True)
        Next
        Me.numMiles.Value = 0
        Me.txtZipCity.Text = ""
        Me.chlstZipCity.Items.Clear()
        Me.chkWC.Checked = False
        Me.chPC.Checked = False
        Me.chRehash.Checked = False
        Me.chlstOrderBy.Items.Remove("Reference Rating")
        Me.chlstOrderBy.Items.Remove("Cash")
        Me.chlstOrderBy.Items.Remove("Finance")
        Me.chlstOrderBy.Items.Remove("Sale Closed Date")
        Me.chlstOrderBy.Items.Remove("Loan Satisfied Date")
        Me.chlstOrderBy.Items.Remove("Approved Loan Amount")
       
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()

    End Sub

 


    Private Sub tabDate_Leave(sender As Object, e As EventArgs) Handles tabDate.Leave
        Me.chlstMarketers.Items.Clear()
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand
        Dim r1 As SqlDataReader
        If Me.chWeekdays.Checked = True Then
            cmdGet = New SqlCommand("Select Distinct(Marketer) from enterlead where ApptDate Between '" & Me.dpFrom.Value.ToString & "' and '" & Me.dpTo.Value.ToString & "' and ApptTime Between '" & Me.tpFrom.Value.ToString & "' and '" & Me.tpTo.Value.ToString & "' and LeadGeneratedOn >= '" & Me.dpGenerated.Value.ToString & "' and (ApptDay <> 'Saturday' and ApptDay <> 'Sunday')", cnn)
        Else
            cmdGet = New SqlCommand("Select Distinct(Marketer) from enterlead where ApptDate Between '" & Me.dpFrom.Value.ToString & "' and '" & Me.dpTo.Value.ToString & "' and ApptTime Between '" & Me.tpFrom.Value.ToString & "' and '" & Me.tpTo.Value.ToString & "' and LeadGeneratedOn >= '" & Me.dpGenerated.Value.ToString & "'", cnn)
        End If

        cmdGet.CommandType = CommandType.Text
        cnn.Open()
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        While r1.Read()
            Me.chlstMarketers.Items.Add(r1.Item(0))
        End While
        For x As Integer = 0 To Me.chlstMarketers.Items.Count - 1
            Me.chlstMarketers.SetItemChecked(x, True)
        Next
    End Sub

    Private Sub checkMarketers_Click(sender As Object, e As EventArgs) Handles checkMarketers.Click
        For x As Integer = 0 To Me.chlstMarketers.Items.Count - 1
            Me.chlstMarketers.SetItemChecked(x, True)
        Next
    End Sub

    Private Sub uncheckMarketers_Click(sender As Object, e As EventArgs) Handles uncheckMarketers.Click
        For x As Integer = 0 To Me.chlstMarketers.Items.Count - 1
            Me.chlstMarketers.SetItemChecked(x, False)
        Next
    End Sub

    Private Sub rdoCash_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCash.CheckedChanged
        If Me.rdoCash.Checked = True Then
            ''Lock Financing part of PC query
            Me.chLoanSatisfied.CheckState = False
            Me.chLoanSatisfied.Enabled = False
            Me.numMonths.Value = 0
            Me.numMonths.Enabled = False
            Me.chApprovedFor.Checked = False
            Me.chApprovedFor.Enabled = False
            Me.txtApprovedDollars.Text = ""
            Me.txtApprovedDollars.Enabled = False
        Else
            Me.chLoanSatisfied.Enabled = True
            Me.numMonths.Enabled = True
            Me.chApprovedFor.Enabled = True
            Me.txtApprovedDollars.Enabled = True

        End If
    End Sub

     

    Private Sub tabProducts_Leave(sender As Object, e As EventArgs) Handles tabProducts.Leave
        If Me.chlstProducts.Items.Count <> Me.chlstProducts.SelectedItems.Count Then
            Me.chFutureInterest.Checked = True
        End If
    End Sub


    Private Sub btnCheckOrderBy_Click(sender As Object, e As EventArgs) Handles btnCheckOrderBy.Click
        For x As Integer = 0 To Me.chlstOrderBy.Items.Count - 1
            Me.chlstOrderBy.SetItemChecked(x, True)
        Next
    End Sub

    Private Sub btnUnCheckOrderBy_Click(sender As Object, e As EventArgs) Handles btnUnCheckOrderBy.Click
        For x As Integer = 0 To Me.chlstOrderBy.Items.Count - 1
            Me.chlstOrderBy.SetItemChecked(x, False)
        Next
    End Sub

    Private Sub btnUp_Click(sender As Object, e As EventArgs) Handles btnUp.Click
        If Me.chlstOrderBy.SelectedItems.Count = 0 Then
            Exit Sub
        End If
        If Me.chlstOrderBy.SelectedIndex = 0 Then
            Exit Sub
        End If
        If Me.chlstOrderBy.SelectedIndex > 0 Then
            Dim Index As Integer = Me.chlstOrderBy.Items.IndexOf(Me.chlstOrderBy.SelectedItem)
            Dim What As String = Me.chlstOrderBy.SelectedItem.ToString
            Dim ChkState As Boolean = Me.chlstOrderBy.GetItemCheckState(Index)
            Me.chlstOrderBy.Items.RemoveAt(Index)
            Me.chlstOrderBy.Items.Insert(Index - 1, What)
            Me.chlstOrderBy.SelectedItem = What
            If ChkState = True Then
                Me.chlstOrderBy.SetItemCheckState(Index - 1, CheckState.Checked)
            Else
                Me.chlstOrderBy.SetItemCheckState(Index - 1, CheckState.Unchecked)
            End If
        End If
    End Sub

    Private Sub btnDown_Click(sender As Object, e As EventArgs) Handles btnDown.Click
        If Me.chlstOrderBy.SelectedItems.Count = 0 Then
            Exit Sub
        End If
        If Me.chlstOrderBy.SelectedIndex = Me.chlstOrderBy.Items.Count - 1 Then
            Exit Sub
        End If
        If Me.chlstOrderBy.SelectedIndex < Me.chlstOrderBy.Items.Count - 1 Then
            Dim Index As Integer = Me.chlstOrderBy.Items.IndexOf(Me.chlstOrderBy.SelectedItem)
            Dim What As String = Me.chlstOrderBy.SelectedItem.ToString
            Dim ChkState As Boolean = Me.chlstOrderBy.GetItemCheckState(Index)
            Me.chlstOrderBy.Items.RemoveAt(Index)
            Me.chlstOrderBy.Items.Insert(Index + 1, What)
            Me.chlstOrderBy.SelectedItem = What
            If ChkState = True Then
                Me.chlstOrderBy.SetItemCheckState(Index + 1, CheckState.Checked)
            Else
                Me.chlstOrderBy.SetItemCheckState(Index + 1, CheckState.Unchecked)
            End If


        End If
    End Sub

    Private Sub tabPC_Leave(sender As Object, e As EventArgs) Handles tabPC.Leave
        Me.chlstOrderBy.Items.Remove("Reference Rating")
        Me.chlstOrderBy.Items.Remove("Cash")
        Me.chlstOrderBy.Items.Remove("Finance")
        Me.chlstOrderBy.Items.Remove("Sale Closed Date")
        Me.chlstOrderBy.Items.Remove("Loan Satisfied Date")
        Me.chlstOrderBy.Items.Remove("Approved Loan Amount")
        If Me.chPC.Checked = True Then
            Me.chlstOrderBy.Items.Add("Reference Rating")
            Me.chlstOrderBy.Items.Add("Cash")
            Me.chlstOrderBy.Items.Add("Finance")
            Me.chlstOrderBy.Items.Add("Sale Closed Date")
            If Me.chLoanSatisfied.Checked = True Then
                Me.chlstOrderBy.Items.Add("Loan Satisfied Date")
            End If
            If Me.chApprovedFor.Checked = True Then
                Me.chlstOrderBy.Items.Add("Approved Loan Amount")
            End If
        End If
    End Sub
End Class
