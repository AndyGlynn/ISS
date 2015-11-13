Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class WCaller
    Friend WithEvents btnUndoSet As New ToolStripButton
    Friend WithEvents btnUndoMemorize As New ToolStripMenuItem
    Friend WithEvents btnAlt1 As New ToolStripMenuItem
    Friend WithEvents btnAlt2 As New ToolStripMenuItem
    Friend WithEvents btnMain As New ToolStripMenuItem
    Friend WithEvents separator As New ToolStripSeparator
    Friend WithEvents tsbtnAlt1 As New ToolStripMenuItem
    Friend WithEvents tsbtnAlt2 As New ToolStripMenuItem
    Friend WithEvents rcbtnSet As New ToolStripMenuItem
    Friend WithEvents rcbtnShow As New ToolStripMenuItem
    Friend WithEvents rcbtnUndoSet As New ToolStripMenuItem
    Friend WithEvents rcbtnRemoveMemorized As New ToolStripMenuItem
    Friend WithEvents rcbtnMemorize As New ToolStripMenuItem
    Public LoadComplete As Boolean = False
    Public LastD1 As String
    Public LastD2 As String
    Public Tab As String = "WC"
    Public PBar As Boolean = False
    Public msgzip As Boolean = False
    Public msgcity As Boolean = False
    Public cntZip As Integer
    Public cntCity As Integer
    Public Toolbar As Integer = 1
    Public TT As ToolTip
    Public SelItem As New ListViewItem
    Public xcordinate As Integer = 0
    Public ycordinate As Integer = 0
    Public IfExists As Boolean = False
    Public ID As String
    Private Sub WCaller_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub WCaller_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim c As New WarmCalling.LoadProcedure()

    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        If Me.TabControl2.SelectedIndex = (1) Then
            Me.Tab = "MA"
            Me.btnExpandCollapse.Enabled = False
        If Me.lvMyAppts.Items.Count > 0 And Me.lvMyAppts.SelectedItems.Count = 0 Then
                Me.lvMyAppts.TopItem.Selected = True
            End If
            If Me.lvMyAppts.SelectedItems.Count = 0 Or Me.lvMyAppts.Items.Count = 0 Then
                Me.ToolBarConfig(2)
            Else
                Me.lvMyAppts_SelectedIndexChanged(Nothing, Nothing)
            End If
            Dim w = Me.Size.Width
            Me.pnlSearch.Controls.Remove(Me.btnExpandCollapse)
            Me.Controls.Add(Me.btnExpandCollapse)

            Me.btnExpandCollapse.Location = New System.Drawing.Point(w - 30, 28)

            Me.btnExpandCollapse.Text = Chr(171)

            Me.pnlSearch.Visible = False
            Me.btnExpandCollapse.BringToFront()
        Else
            If Me.IfExists = True Then
                Me.TT.Dispose()
            End If
            If Me.lvWarmCalling.Items.Count > 0 And Me.lvWarmCalling.SelectedItems.Count = 0 Then
                Me.lvWarmCalling.TopItem.Selected = True
            End If
            Me.Tab = "WC"
            Me.ToolBarConfig(1)
            Me.lvWarmCalling_SelectedIndexChanged(Nothing, Nothing)
            Me.btnExpandCollapse.Enabled = True
        End If
    End Sub

    Private Sub lvMyAppts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMyAppts.Click
        If Me.IfExists = True Then
            Me.TT.Dispose()
        End If
        Me.SelItem.ToolTipText = Nothing
        Me.SelItem.ToolTipText = Me.lvMyAppts.GetItemAt(Me.xcordinate, Me.ycordinate).ToolTipText
        'MsgBox(Me.SelItem.Tag)
    End Sub

    Private Sub lvMyAppts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMyAppts.DoubleClick
        Me.TT = New ToolTip
        Me.IfExists = True
        If Me.SelItem Is Nothing Then
            Exit Sub
        End If
        Dim p As System.Drawing.Point = New System.Drawing.Point(10, Me.ycordinate + 15)
        Dim notes As String
        If Me.SelItem.ToolTipText <> "" Then
            Dim c As New TruncateNotes
            c.Truncate(Me.SelItem.ToolTipText, Me.lvMyAppts)
            notes = c.NewSTRING
            TT.ToolTipIcon = ToolTipIcon.Info
            Dim z = Me.lvMyAppts.SelectedItems(0).Index
            Dim y As ListViewGroup = Me.lvMyAppts.Items(z).Group
            If y.Name <> ("grpMemorized") Then
                TT.ToolTipTitle = "Set Appt. Notes"
            Else
                TT.ToolTipTitle = "Memorized Appt. Notes"
            End If
            Me.TT.Show(notes, Me.lvMyAppts, p, 30000)
        End If
    End Sub

    Private Sub lvMyAppts_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvMyAppts.MouseMove
        Me.xcordinate = e.X
        Me.ycordinate = e.Y

    End Sub

    Private Sub lvMyAppts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMyAppts.SelectedIndexChanged
        If Me.Tab = "MA" Then
            If Me.lvMyAppts.SelectedItems.Count = 0 Or Me.lvMyAppts.Items.Count = 0 Then
                Me.ToolBarConfig(2)
                Exit Sub
            End If
            Dim z = Me.lvMyAppts.SelectedItems(0).Index
            Dim y As ListViewGroup = Me.lvMyAppts.Items(z).Group
            If y.Name <> ("grpMemorized") Then
                Me.ToolBarConfig(3)
            Else
                Me.ToolBarConfig(2)
            End If
            Dim c As New WarmCalling
            If Me.lvMyAppts.SelectedItems.Count = 0 Then
                c.PullCustomerINFO("")
            Else
                c.PullCustomerINFO(Me.lvMyAppts.SelectedItems(0).Tag)

                Me.btnAutoDialer.DropDownItems.Add(Me.separator)
                Me.btnAutoDialer.DropDownItems.Add(Me.btnMain)
                Me.btnMain.Text = "Call Main- " & Me.txtHousePhone.Text
                If Me.txtaltphone1.Text <> "" Then
                    Me.btnAutoDialer.DropDownItems.Add(Me.btnAlt1)
                    Me.btnAlt1.Text = "Call Alt 1- " & Me.txtaltphone1.Text
                Else
                    Me.btnAutoDialer.DropDownItems.Remove(Me.btnAlt1)
                End If
                If Me.txtaltphone2.Text <> "" Then
                    Me.btnAutoDialer.DropDownItems.Add(Me.btnAlt2)
                    Me.btnAlt2.Text = "Call Alt 2- " & Me.txtaltphone2.Text
                Else
                    Me.btnAutoDialer.DropDownItems.Remove(Me.btnAlt2)
                End If
            End If
        End If
    End Sub

    Private Sub ToolBarConfig(ByVal x As Integer)
        Select Case x
            Case 1
                Me.tsWarmCalling.Items.Clear()
                Me.tsWarmCalling.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnSetAppt, Me.btnEditCustomer, Me.btnChangeStatus, Me.lblTo, Me.lblFrom, Me.cboDateRange, Me.lblDateRange, Me.btnAutoDialer})
                Me.btnSetAppt.DropDownItems.Clear()
                Me.btnSetAppt.DropDownItems.Add(Me.MemorizeThisApptToolStripMenuItem)
                Me.dtp1.Visible = True
                Me.dtp2.Visible = True
                Me.lblDateRange.Visible = True
                Me.lblFrom.Visible = True
                Me.lblTo.Visible = True
                If Me.txtDate1.Text = "" Then
                    Me.txtDate1.Visible = True
                End If
                If Me.txtDate2.Text = "" Then
                    Me.txtDate2.Visible = True
                End If
                If Me.ContextMenuStrip1.Items.Count <> 0 Then
                    Me.ContextMenuStrip1.Items.Clear()
                End If


                Me.ContextMenuStrip1.Items.Add("Set Appointment", Me.ilToolStripIcons.Images(3), AddressOf btnSetAppt_ButtonClick)
                Me.ContextMenuStrip1.Items.Add("Memorize Appointment", Me.ilToolStripIcons.Images(4), AddressOf MemorizeThisApptToolStripMenuItem_Click)
            Case 2
                Me.tsWarmCalling.Items.Clear()
                Me.dtp1.Visible = False
                Me.dtp2.Visible = False
                Me.txtDate1.Visible = False
                Me.txtDate2.Visible = False
                Me.tsWarmCalling.Items.Clear()
                Me.tsWarmCalling.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnSetAppt, Me.btnEditCustomer, Me.btnChangeStatus, Me.btnAutoDialer})
                Me.btnSetAppt.DropDownItems.Clear()
                Me.btnSetAppt.DropDownItems.Add(Me.btnUndoMemorize)
                If Me.ContextMenuStrip1.Items.Count <> 0 Then
                    Me.ContextMenuStrip1.Items.Clear()
                End If
                Me.ContextMenuStrip1.Items.Add("Show Notes", Me.ilToolStripIcons.Images(2), AddressOf Me.lvMyAppts_DoubleClick)
                Me.ContextMenuStrip1.Items.Add("Set Appointment", Me.ilToolStripIcons.Images(3), AddressOf btnSetAppt_ButtonClick)
                Me.ContextMenuStrip1.Items.Add("Remove Memorized Appt.", Me.ilToolStripIcons.Images(1), AddressOf Me.btnUndoMemorize_Click)
            Case 3
                Me.tsWarmCalling.Items.Clear()
                Me.dtp1.Visible = False
                Me.dtp2.Visible = False
                Me.txtDate1.Visible = False
                Me.txtDate2.Visible = False
                Me.tsWarmCalling.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnUndoSet, Me.btnEditCustomer, Me.btnChangeStatus, Me.btnAutoDialer})
                If Me.ContextMenuStrip1.Items.Count <> 0 Then
                    Me.ContextMenuStrip1.Items.Clear()
                End If
                Me.ContextMenuStrip1.Items.Add("Show Notes", Me.ilToolStripIcons.Images(2), AddressOf Me.lvMyAppts_DoubleClick)

                Me.ContextMenuStrip1.Items.Add("Undo Set Appt.", Me.ilToolStripIcons.Images(1), AddressOf Me.btnUndoSet_Click)
        End Select
    End Sub

    Private Sub lvWarmCalling_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvWarmCalling.SelectedIndexChanged
        If Me.Tab = "WC" Then
            Dim c As New WarmCalling
            Me.ToolBarConfig(1)
            If Me.lvWarmCalling.SelectedItems.Count > 0 Then
                c.PullCustomerINFO(Me.lvWarmCalling.SelectedItems(0).Text)

                Me.btnAutoDialer.DropDownItems.Add(Me.separator)
                Me.btnAutoDialer.DropDownItems.Add(Me.btnMain)
                Me.btnMain.Text = "Call Main- " & Me.txtHousePhone.Text
                If Me.txtaltphone1.Text <> "" Then
                    Me.btnAutoDialer.DropDownItems.Add(Me.btnAlt1)
                    Me.btnAlt1.Text = "Call Alt 1- " & Me.txtaltphone1.Text
                Else
                    Me.btnAutoDialer.DropDownItems.Remove(Me.btnAlt1)

                End If
                If Me.txtaltphone2.Text <> "" Then
                    Me.btnAutoDialer.DropDownItems.Add(Me.btnAlt2)
                    Me.btnAlt2.Text = "Call Alt 2- " & Me.txtaltphone2.Text
                Else
                    Me.btnAutoDialer.DropDownItems.Remove(Me.btnAlt2)
                End If
            Else
                c.PullCustomerINFO("")
            End If
        End If
    End Sub

    Private Sub btnExpandWarmCalling_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExpandWarmCalling.Click
        Select Case Me.btnExpandWarmCalling.Text
            Case Chr(187)
                Me.SplitContainer1.SplitterDistance = 2500
                Me.SplitContainer1.SplitterWidth = 100
                Me.btnExpandMyAppts.Text = Chr(171)
                Me.btnExpandWarmCalling.Text = Chr(171)
            Case Chr(171)
                Me.SplitContainer1.SplitterDistance = 225
                Me.SplitContainer1.SplitterWidth = 5
                Me.btnExpandWarmCalling.Text = Chr(187)
                Me.btnExpandMyAppts.Text = Chr(187)
        End Select
    End Sub

    Private Sub btnExpandMyAppts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExpandMyAppts.Click
        Me.btnExpandWarmCalling_Click(Nothing, Nothing)
    End Sub

    Private Sub cboPLSWarmCalling_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPLSWarmCalling.GotFocus
        Me.lblPLSWarmCalling.Visible = False
    End Sub

    Private Sub cboPLSWarmCalling_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPLSWarmCalling.LostFocus
        If Me.cboPLSWarmCalling.Text = "" Then
            Me.lblPLSWarmCalling.Visible = True
        End If
    End Sub

    Private Sub cboSLSWarmCalling_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSLSWarmCalling.GotFocus
        Me.lblSLSWarmCalling.Visible = False
        If Me.cboPLSWarmCalling.Text = "" Then
            Me.cboSLSWarmCalling.Items.Clear()
        End If

    End Sub

    Private Sub cboSLSWarmCalling_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSLSWarmCalling.LostFocus
        If Me.cboSLSWarmCalling.Text = "" Then
            Me.lblSLSWarmCalling.Visible = True
        End If
    End Sub

    Private Sub lblPLSWarmCalling_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPLSWarmCalling.Click
        Me.cboPLSWarmCalling.Focus()
    End Sub

    Private Sub lblSLSWarmCalling_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblSLSWarmCalling.Click
        Me.cboSLSWarmCalling.Focus()
    End Sub

    Private Sub btnExpandCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandCollapse.Click
        Dim w = Me.Size.Width
        If Me.btnExpandCollapse.Text = Chr(171) Then
            Me.Controls.Remove(Me.btnExpandCollapse)
            Me.pnlSearch.Controls.Add(Me.btnExpandCollapse)
            Me.btnExpandCollapse.Location = New System.Drawing.Point(3, 3)
            Me.btnExpandCollapse.Text = Chr(187)
            Me.pnlSearch.Location = New System.Drawing.Point(w - 344, 28)
            Me.pnlSearch.Size = New System.Drawing.Size(344, 576)
            Me.pnlSearch.Visible = True
            Me.btnExpandCollapse.BringToFront()
            Dim x As Integer = Me.pnlSearch.Size.Height

            Me.btnExpandCollapse.Size = New Size(Me.btnExpandCollapse.Size.Width, x - 6)

        Else

            Me.pnlSearch.Controls.Remove(Me.btnExpandCollapse)
            Me.Controls.Add(Me.btnExpandCollapse)

            Me.btnExpandCollapse.Location = New System.Drawing.Point(w - 30, 31)

            Me.btnExpandCollapse.Text = Chr(171)

            Me.pnlSearch.Visible = False
            Me.btnExpandCollapse.BringToFront()
        End If
    End Sub

    Private Sub rdoCity_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoCity.CheckedChanged
        Me.lblEnter.Text = "Enter Starting City:"
        Me.lblShow.Text = "Show Cities within                      miles of" & vbCrLf & "starting City"
        Me.btnZipCity.Text = "Show Cities"
        Dim c As New AutoCompleteSourceCities
        Me.txtZipCode.Text = ""
    End Sub

    Private Sub rdoZip_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoZip.CheckedChanged
        Me.lblEnter.Text = "Enter Starting Zip Code:"
        Me.lblShow.Text = "Show Zip Codes within                miles of" & vbCrLf & "starting Zip Code"
        Me.btnZipCity.Text = "Show Zip Codes"
        Me.txtZipCode.AutoCompleteCustomSource.Clear()
        Me.txtZipCode.Text = ""
    End Sub


    Private Sub WCaller_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        Dim w = Me.Size.Width
        If Me.btnExpandCollapse.Parent.Name = "pnlSearch" Then
            If Me.pnlSearch.Visible = True Then
                If Me.Size.Width = 998 And Me.Size.Height = 635 Then
                    Me.btnExpandCollapse.Size = New Size(21, 570)
                End If
            End If
            Exit Sub
        Else
            Me.btnExpandCollapse.Location = New System.Drawing.Point(w - 30, 31)
        End If
        Dim c As New CustomerHistory
        c.SetUp(Me, Me.ID, Me.TScboCustomerHistory)

    End Sub


    Private Sub LaunchProgressiveDialerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaunchProgressiveDialerToolStripMenuItem.Click
        Me.tsAutoDial.Visible = True
    End Sub

    Private Sub tsbtnCloseDialer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtnCloseDialer.Click
        Me.tsAutoDial.Visible = False
    End Sub

    Private Sub btnLogCall_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogCall.ButtonClick

        LogPhoneCall.frm = Me
        If Me.Tab = "WC" Then
            If Me.lvWarmCalling.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            End If
            LogPhoneCall.ID = Me.lvWarmCalling.SelectedItems(0).Text
        Else
            If Me.lvMyAppts.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            End If
            LogPhoneCall.ID = Me.lvMyAppts.SelectedItems(0).Tag
        End If
        LogPhoneCall.Contact1 = Me.txtContact1.Text
        LogPhoneCall.Contact2 = Me.txtContact2.Text
        LogPhoneCall.ShowInTaskbar = False
        LogPhoneCall.ShowDialog()
    End Sub

    Private Sub LogAsCalledCancelledToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogAsCalledCancelledToolStripMenuItem.Click

        Dim s = Split(Me.txtContact1.Text, " ")
        Dim c1 = s(0)
        Dim s2 = Split(Me.txtContact2.Text, " ")
        Dim c2 = s2(0)
        CandCNotes.Contact1 = c1
        CandCNotes.Contact2 = c2
        CandCNotes.OrigApptDate = Me.txtApptDate.Text
        CandCNotes.OrigApptTime = Me.txtApptTime.Text

        CandCNotes.frm = Me
        If Me.Tab = "WC" And Me.lvWarmCalling.SelectedItems.Count <> 0 Then
            CandCNotes.ID = Me.lvWarmCalling.SelectedItems(0).Text
        ElseIf Me.Tab = "MA" And Me.lvMyAppts.SelectedItems.Count <> 0 Then
            CandCNotes.ID = Me.lvMyAppts.SelectedItems(0).Tag
        Else
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        CandCNotes.ShowInTaskbar = False
        CandCNotes.ShowDialog()
        'Stamp in Log without notes

    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.btnClear.Enabled = False
        Me.txtTime1.Text = ""
        Me.txtTime1.Visible = True
        Me.txtTime2.Text = ""
        Me.txtTime2.Visible = True
        Me.cbWeekdays.Checked = False
        For i As Integer = 0 To Me.chlstResults.Items.Count - 1
            Me.chlstResults.SetItemChecked(i, True)
        Next
        Me.txtZipCode.Text = ""
        Me.nupMiles.Value = 0
        Me.lbZipCity.Items.Clear()
        Dim c As New WarmCalling
        c.ManagerCriteria()
        c.Populate()
        Me.btnClear.Enabled = True
    End Sub

    Public Sub lblCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCheckAll.Click


        For i As Integer = 0 To Me.lbZipCity.Items.Count - 1
            Me.lbZipCity.SetItemChecked(i, True)
        Next


    End Sub

    Private Sub lblUncheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblUncheckAll.Click
        For i As Integer = 0 To Me.lbZipCity.Items.Count - 1
            Me.lbZipCity.SetItemChecked(i, False)
        Next
    End Sub

    Private Sub btnUndoMemorize_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUndoMemorize.Click
        If Me.lvMyAppts.SelectedItems.Count = 0 Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If

        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.RemoveMemorized", cnn)
        Dim r1 As SqlDataReader
        Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param2 As SqlParameter = New SqlParameter("@Form", "Warm Calling")
        Dim param3 As SqlParameter = New SqlParameter("@ID", Me.lvMyAppts.SelectedItems(0).Tag)

        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)

        cmdGet.CommandType = CommandType.StoredProcedure
        cnn.Open()
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r1.Read()
        r1.Close()
        cnn.Close()

        Me.lvMyAppts.SelectedItems(0).Remove()
        If Me.lvMyAppts.Items.Count <> 0 Then
            Me.lvMyAppts.TopItem.Selected = True
        End If


    End Sub

    Private Sub cboGroupBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupBy.LostFocus
        If Me.cboGroupBy.Text = "" Then
            Me.lblGroupBy.Visible = True
        End If
    End Sub

    Private Sub cboGroupBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupBy.GotFocus
        Me.lblGroupBy.Visible = False
    End Sub

    Private Sub lblGroupBy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblGroupBy.Click
        Me.cboGroupBy.Focus()
    End Sub

    Private Sub txtTime1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime1.Click
        Me.txtTime1.Visible = False
        Me.dtpTime1.Focus()
        Me.txtTime1.Text = Me.dtpTime1.Value.ToShortTimeString
    End Sub

    Private Sub txtTime2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime2.Click
        Me.txtTime2.Visible = False
        Me.dptTime2.Focus()
        Me.txtTime2.Text = Me.dptTime2.Value.ToShortTimeString
    End Sub

    Private Sub dtpTime1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTime1.GotFocus
        Me.txtTime1.Text = Me.dtpTime1.Value.ToShortTimeString
    End Sub

    Private Sub dtpTime1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTime1.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtTime1.Text = Me.dtpTime1.Value.ToShortTimeString
    End Sub

    Private Sub dptTime2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dptTime2.GotFocus

        Me.txtTime2.Text = Me.dptTime2.Value.ToShortTimeString
    End Sub

    Private Sub dptTime2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dptTime2.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtTime2.Text = Me.dptTime2.Value.ToShortTimeString
    End Sub

    Private Sub txtDate2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate2.Click
        Me.dtp2.Focus()
        Me.txtDate2.Visible = False

        Me.txtDate2.Text = Me.dtp2.Value.ToShortDateString
        Me.cboDateRange.Text = "Custom"
    End Sub

    Private Sub txtDate1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate1.Click
        Me.txtDate1.Visible = False
        Me.dtp1.Focus()
        Me.txtDate1.Text = Me.dtp1.Value.ToShortDateString
        Me.cboDateRange.Text = "Custom"
    End Sub

    Private Sub dtp1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp1.GotFocus
        Me.txtDate1.Visible = False
        Me.txtDate1.Text = Me.dtp1.Value.ToShortDateString
        'Me.txtDate2.Visible = False
        Me.cboDateRange.Text = "Custom"
    End Sub

    Private Sub dtp1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp1.LostFocus
        If Me.txtDate1.Text = Me.LastD1 And Me.txtDate2.Text = Me.LastD2 Then
            Exit Sub
        End If
        Dim c As New WarmCalling
        c.Populate()
    End Sub

    Private Sub dtp1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp1.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtDate1.Text = Me.dtp1.Value.ToShortDateString

    End Sub

    Private Sub dtp2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp2.GotFocus
        'Me.txtDate1.Visible = False
        Me.txtDate2.Text = Me.dtp2.Value.ToShortDateString
        Me.txtDate2.Visible = False
        Me.cboDateRange.Text = "Custom"
    End Sub

    Private Sub dtp2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp2.LostFocus
        Dim c As New WarmCalling
        'If Me.txtDate1.Text = "" Then
        '    MsgBox("You must supply a ""Start Date""", MsgBoxStyle.Exclamation, "Please Supply a Date")
        '    Me.dtp1.Focus()
        '    Exit Sub
        'End If
        If Me.LastD1 = Me.txtDate1.Text And Me.LastD2 = Me.txtDate2.Text Then
            Exit Sub
        End If
        c.Populate()



    End Sub

    Private Sub dtp2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp2.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Me.txtDate2.Text = Me.dtp2.Value.ToShortDateString

    End Sub

    Private Sub cboDateRange_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDateRange.SelectedIndexChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        If Me.cboDateRange.Text = "All" Then
            Me.txtDate1.Text = ""
            Me.txtDate1.Visible = True
            Me.txtDate2.Text = ""
            Me.txtDate2.Visible = True

        End If

        If Me.cboDateRange.Text <> "All" And Me.cboDateRange.Text <> "Custom" Then
            Dim d As New DTPManipulation(Me.cboDateRange.Text)
            Me.dtp1.Value = d.retDateFrom
            Me.dtp2.Value = d.retDateTo
            Me.txtDate1.Visible = False
            Me.txtDate2.Visible = False
            Me.txtDate1.Text = d.retDateFrom.ToString
            Me.txtDate2.Text = d.retDateTo.ToString
        End If
        If Me.LastD1 = Me.txtDate1.Text And Me.LastD2 = Me.txtDate2.Text Then
            Exit Sub
        End If
        Dim c As New WarmCalling
        c.Populate()
    End Sub

    Public Sub btnZipCity_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnZipCity.Click
        Me.pbRadiusSearch.Value = 0
        Me.pbRadiusSearch.Visible = True
        Me.btnZipCity.Enabled = False
        If Me.txtZipCode.Text = "" Or Me.nupMiles.Value = 0 Then
            MsgBox("You must supply a starting point and" & vbCr & "supply a radius of at least 1 mile!", MsgBoxStyle.Exclamation, "Please supply Values")
            Me.btnZipCity.Enabled = True
            Exit Sub

        End If

        Dim c As New MAPPOINT_LOGIC
        If Me.rdoZip.Checked = True Then
            c.DoIt(Me.nupMiles.Value, Me.txtZipCode.Text)
        Else
            Dim b As New GetStateFromCity(Me.txtZipCode.Text)
            c.DoItCity(Me.nupMiles.Value, Me.txtZipCode.Text, b.StatePulled)
        End If
        Me.btnZipCity.Enabled = True
    End Sub

    Private Sub cboGroupBy_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupBy.SelectedValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Dim c As New WarmCalling
        If Me.cboGroupBy.Text = "" Then
            Me.lvWarmCalling.Groups.Clear()
            c.Populate()
            Exit Sub
        End If
        c.GroupBy()
    End Sub

    Private Sub cboPLSWarmCalling_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPLSWarmCalling.SelectedIndexChanged
        Dim c As New WarmCalling

        If Me.LoadComplete = False Then
            Exit Sub
        End If
        If Me.cboPLSWarmCalling.Text <> "" Then
            c.GetSLS(Me.cboPLSWarmCalling.Text)
        End If

        c.Populate()
    End Sub

    Private Sub cboSLSWarmCalling_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSLSWarmCalling.SelectedIndexChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        Dim c As New WarmCalling
        c.Populate()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        If Me.lbZipCity.Items.Count > 0 And Me.lbZipCity.CheckedItems.Count = 0 Then
            MsgBox("You must check at least 1 Zip/City to Search", MsgBoxStyle.Exclamation, "No Zip/City Selected")
            Me.lbZipCity.SetSelected(0, True)
            Me.lbZipCity.SetItemChecked(0, True)
            Exit Sub
        End If
        If Me.chlstResults.CheckedItems.Count = 0 Then
            MsgBox("You must check at least 1 marketing result to search", MsgBoxStyle.Exclamation, "No Marketing Result Selected")
            Exit Sub
        End If
        Dim c As New WarmCalling
        c.Populate()
    End Sub

    Private Sub btnKill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKill.Click
        If Me.Tab = "WC" Then
            If Me.lvWarmCalling.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            End If
        Else
            If Me.lvMyAppts.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            End If
        End If


        Kill.Contact1 = Me.txtContact1.Text
        Kill.Contact2 = Me.txtContact2.Text
        Kill.frm = "WC"
        If Me.Tab = "WC" Then
            Kill.ID = Me.lvWarmCalling.SelectedItems(0).Text
        Else
            Kill.ID = Me.lvMyAppts.SelectedItems(0).Tag
        End If
        Kill.ShowInTaskbar = False
        Kill.ShowDialog()
    End Sub

    Private Sub btnDoNotCall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDoNotCall.Click
        If Me.Tab = "WC" Then
            If Me.lvWarmCalling.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            Else
                Dim c As New DoNotCallOrMail
                c.DoNot(Me.lvWarmCalling.SelectedItems(0).Text, sender.text.ToString)
            End If
            Dim i As Integer = Me.lvWarmCalling.Items.IndexOfKey(Me.lvWarmCalling.SelectedItems(0).Text)
            Dim x As New WarmCalling.MyApptsTab.Populate(Me.cboFilter.Text)
            Me.lvWarmCalling.SelectedItems(0).Remove()
            Me.txtRecordsMatching.Text = CStr(CInt(Me.txtRecordsMatching.Text) - 1)
            If Me.lvWarmCalling.Items.Count <> 0 Then
                If i > Me.lvWarmCalling.Items.Count - 1 Then
                    Me.lvWarmCalling.Items(i - 1).Selected = True
                Else
                    Me.lvWarmCalling.Items(i).Selected = True
                End If
            Else
                Dim wc As New WarmCalling
                wc.PullCustomerINFO("")
            End If
        Else
            If Me.lvMyAppts.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            Else
                Dim c As New DoNotCallOrMail
                c.DoNot(Me.lvMyAppts.SelectedItems(0).Tag, sender.text.ToString)
                Dim y As New WarmCalling
                y.Populate()
                Dim x As New WarmCalling.MyApptsTab.Populate(Me.cboFilter.Text)
            End If
        End If

    End Sub

    Private Sub btnDoNotMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDoNotMail.Click
        If Me.Tab = "WC" Then

            If Me.lvWarmCalling.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            Else
                Dim y As New DoNotCallOrMail
                y.DoNot(Me.lvWarmCalling.SelectedItems(0).Text, sender.text.ToString)
            End If
            Dim c As New CustomerHistory
            c.SetUp(Me, Me.lvWarmCalling.SelectedItems(0).Text, Me.TScboCustomerHistory)
        Else
            If Me.lvMyAppts.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            Else
                Dim x As New DoNotCallOrMail
                x.DoNot(Me.lvMyAppts.SelectedItems(0).Tag, sender.text.ToString)
            End If
            Dim c As New CustomerHistory
            c.SetUp(Me, Me.lvMyAppts.SelectedItems(0).Tag, Me.TScboCustomerHistory)
        End If

    End Sub

    Private Sub btnDoNotCallOrMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDoNotCallOrMail.Click
        If Me.Tab = "WC" Then
            If Me.lvWarmCalling.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            Else
                Dim c As New DoNotCallOrMail
                c.DoNot(Me.lvWarmCalling.SelectedItems(0).Text, sender.text.ToString)
            End If
            Dim x As New WarmCalling.MyApptsTab.Populate(Me.cboFilter.Text)
            Dim i As Integer = Me.lvWarmCalling.Items.IndexOfKey(Me.lvWarmCalling.SelectedItems(0).Text)
            Me.lvWarmCalling.SelectedItems(0).Remove()
            Me.txtRecordsMatching.Text = CStr(CInt(Me.txtRecordsMatching.Text) - 1)
            If Me.lvWarmCalling.Items.Count <> 0 Then
                If i > Me.lvWarmCalling.Items.Count - 1 Then
                    Me.lvWarmCalling.Items(i - 1).Selected = True
                Else
                    Me.lvWarmCalling.Items(i).Selected = True
                End If
            Else
                Dim wc As New WarmCalling
                wc.PullCustomerINFO("")
            End If
        Else
            If Me.lvMyAppts.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            Else
                Dim c As New DoNotCallOrMail
                c.DoNot(Me.lvMyAppts.SelectedItems(0).Tag, sender.text.ToString)
                Dim y As New WarmCalling
                y.Populate()
                Dim x As New WarmCalling.MyApptsTab.Populate(Me.cboFilter.Text)

            End If
        End If
    End Sub

    Private Class AutoCompleteSourceCities
        Private cnn As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(static_variables.cnn)
        Public Sub New()
            WCaller.lbZipCity.Items.Clear()
            WCaller.txtZipCode.AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim cmdGET As SqlCommand = New SqlCommand("SELECT City,State from iss.dbo.citypull", cnn)
            Dim r1 As SqlDataReader
            cnn.Open()
            r1 = cmdGET.ExecuteReader
            WCaller.txtZipCode.AutoCompleteCustomSource.Clear()
            While r1.Read
                WCaller.txtZipCode.AutoCompleteCustomSource.Add(r1.Item(0).ToString)
            End While
            r1.Close()
            cnn.Close()
        End Sub
    End Class

    Private Class GetStateFromCity
        Private cnn As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(static_variables.cnn)
        Private ST As String = ""
        Public Property StatePulled() As String
            Get
                Return ST
            End Get
            Set(ByVal value As String)
                ST = value
            End Set
        End Property
        Public Sub New(ByVal City As String)
            Dim cmdGET As SqlCommand = New SqlCommand("SELECT state from iss.dbo.citypull where city = @CTY", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@CTY", City)
            cmdGET.Parameters.Add(param1)
            Dim r1 As SqlDataReader
            cnn.Open()
            r1 = cmdGET.ExecuteReader
            While r1.Read
                Me.StatePulled = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
        End Sub
    End Class


    Private Sub txtZipCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtZipCode.LostFocus
        If IsNumeric(Me.txtZipCode.Text) = True Then
            Exit Sub
        Else
            Dim c As New GetStateFromCity(Me.txtZipCode.Text)
        End If
    End Sub

    Private Sub MemorizeThisApptToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MemorizeThisApptToolStripMenuItem.Click
        If Me.lvWarmCalling.SelectedItems.Count = 0 Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
    

        MemorizeNotes.ShowDialog()

    End Sub

    Private Sub btnEditCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditCustomer.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim x As New WarmCalling.MyApptsTab.Populate(Me.cboGroupBy.Text)

    End Sub

    Private Sub cboDisplayColumn_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDisplayColumn.GotFocus
        Me.lblDisplayColumn.Visible = False
    End Sub

    Private Sub cboDisplayColumn_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDisplayColumn.LostFocus
        If Me.cboDisplayColumn.Text = "" Then
            Me.lblDisplayColumn.Visible = True
        End If
    End Sub

    Private Sub cboDisplayColumn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDisplayColumn.SelectedIndexChanged
        Dim c As New WarmCalling.MyApptsTab.DisplayColumn()
    End Sub

    Private Sub cboFilter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFilter.GotFocus
        Me.lblFilter.Visible = False
    End Sub

    Private Sub cboFilter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFilter.LostFocus
        If Me.cboFilter.Text = "" Then
            Me.lblFilter.Visible = True
        End If
    End Sub

    Private Sub cboFilter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFilter.SelectedIndexChanged
        Dim c As New WarmCalling.MyApptsTab.Populate(Me.cboFilter.Text)
    End Sub

    Private Sub cboGroupSetAppt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupSetAppt.GotFocus
        Me.lblGroupSetAppts.Visible = False
    End Sub

    Private Sub cboGroupSetAppt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupSetAppt.LostFocus
        If Me.cboGroupSetAppt.Text = "" Then
            Me.lblGroupSetAppts.Visible = True
        End If
    End Sub

    Private Sub lblGroupSetAppts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblGroupSetAppts.Click
        Me.cboGroupSetAppt.Focus()
    End Sub

    Private Sub lblDisplayColumn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblDisplayColumn.Click
        Me.cboDisplayColumn.Focus()
    End Sub

    Private Sub lblFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblFilter.Click
        Me.cboFilter.Focus()
    End Sub

    Private Sub cboGroupSetAppt_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGroupSetAppt.SelectedIndexChanged
        Dim c As New WarmCalling.MyApptsTab.Populate(Me.cboFilter.Text)
    End Sub

    Private Sub btnUndoSet_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUndoSet.Click
        'Me.lvMyAppts.SelectedItems(0).Remove()

        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.UndoSet", cnn)
        Dim r1 As SqlDataReader
        Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param2 As SqlParameter = New SqlParameter("@ID", Me.lvMyAppts.SelectedItems(0).Tag)

        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.CommandType = CommandType.StoredProcedure
        cnn.Open()
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r1.Read()
        r1.Close()
        cnn.Close()
        Dim c As New WarmCalling
        Dim y As New WarmCalling.MyApptsTab.Populate(Me.cboFilter.Text)
        c.Populate()

    End Sub

    Private Sub btnSetAppt_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetAppt.ButtonClick
        SetAppt.frm = Me
        If Tab = "WC" Then
            If Me.lvWarmCalling.SelectedItems.Count = 0 Then
                MsgBox("You must select a record to Set Appointment!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            End If
            SetAppt.ID = Me.lvWarmCalling.SelectedItems(0).Text
        Else
            If Me.lvMyAppts.SelectedItems.Count = 0 Then
                MsgBox("You must select a record to Set Appointment!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            End If
            SetAppt.ID = Me.lvMyAppts.SelectedItems(0).Tag
        End If
        Dim s = Split(Me.txtContact1.Text, " ")
        Dim s2 = Split(Me.txtContact2.Text, " ")
        SetAppt.Contact1 = s(0)
        SetAppt.Contact2 = s2(0)
        SetAppt.OrigApptDate = Me.txtApptDate.Text
        SetAppt.OrigApptTime = Me.txtApptTime.Text
        SetAppt.ShowInTaskbar = False
        SetAppt.ShowDialog()
    End Sub

  Private Sub WCaller_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Dim x As Integer = Me.pnlSearch.Size.Height
        If Me.btnExpandCollapse.Parent.Name = "pnlSearch" Then
            Me.btnExpandCollapse.Size = New Size(Me.btnExpandCollapse.Size.Width, x - 6)


        End If
    End Sub


  

End Class
