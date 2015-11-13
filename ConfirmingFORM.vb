
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Public Class Confirming
    Friend WithEvents Confirm As New ToolStripDropDownButton '
    Friend WithEvents EditCustomer As New ToolStripButton '
    Friend WithEvents ChangeStatus As New ToolStripDropDownButton '
    Friend WithEvents UndoConfirm As New ToolStripSplitButton '
    Friend WithEvents Emailc As New ToolStripDropDownButton '
    Friend WithEvents Emailu As New ToolStripDropDownButton
    Friend WithEvents AutoDial As New ToolStripDropDownButton  '
    Friend WithEvents Sales As New ToolStripDropDownButton '
    Friend WithEvents SendNotes As New ToolStripMenuItem
    Friend WithEvents Confirmwith As New ToolStripMenuItem
    Friend WithEvents ConfirmSeparator As New ToolStripSeparator
    Friend WithEvents ConfirmwithContact1 As New ToolStripMenuItem
    Friend WithEvents ConfirmwithContact2 As New ToolStripMenuItem
    Friend WithEvents ConfirmwithBoth As New ToolStripMenuItem
    Friend WithEvents Confirmwithther As New ToolStripMenuItem
    Friend WithEvents Reschedule As New ToolStripMenuItem '
    Friend WithEvents btnKill As New ToolStripMenuItem '
    Friend WithEvents Cancelled As New ToolStripMenuItem
    Friend WithEvents DoNotCall As New ToolStripMenuItem '
    Friend WithEvents DoNotMail As New ToolStripMenuItem '
    Friend WithEvents DoNotCallOrMail As New ToolStripMenuItem '
    Friend WithEvents Correspondence As New ToolStripMenuItem '
    Friend WithEvents MainPhone As New ToolStripMenuItem '
    Friend WithEvents AltPhone1 As New ToolStripMenuItem '
    Friend WithEvents AltPhone2 As New ToolStripMenuItem '
    Friend WithEvents EmailConfirmation1 As New ToolStripMenuItem '
    Friend WithEvents EmailConfirmationAll As New ToolStripMenuItem '
    Friend WithEvents EmailCustomu As New ToolStripMenuItem '
    Friend WithEvents EmailCustomc As New ToolStripMenuItem '
    Friend WithEvents EmailTemplatec As New ToolStripMenuItem
    Friend WithEvents EmailTemplateu As New ToolStripMenuItem
    Friend WithEvents SalesResult As New ToolStripMenuItem
    Friend WithEvents btnCNGApptTime As New ToolStripMenuItem
    Friend WithEvents SalesSeparator As New ToolStripSeparator
    Friend WithEvents Issue As New ToolStripMenuItem
    Friend WithEvents Rep1 As New ToolStripComboBox
    Friend WithEvents Rep2 As New ToolStripComboBox
    Friend WithEvents TemplateListu As New ToolStripComboBox
    Friend WithEvents TemplateListc As New ToolStripComboBox
    Friend WithEvents SaveChanges As New ToolStripMenuItem
    Friend WithEvents TBLabel As New ToolStripLabel
    Friend WithEvents dpConfirming As New DateTimePicker
    Friend WithEvents dpSales As New DateTimePicker
    Dim Seperator As New ToolStripSeparator
    Dim Seperator2 As New ToolStripSeparator
    Dim SCSeparator As New ToolStripSeparator
    Friend WithEvents LaunchAutodial As New ToolStripMenuItem
    Friend WithEvents btnAlt1ad As New ToolStripMenuItem
    Friend WithEvents btnAlt2ad As New ToolStripMenuItem
    Friend WithEvents CreateNewTemplateu As New ToolStripMenuItem
    Friend WithEvents CreateNewTemplatec As New ToolStripMenuItem
    Friend WithEvents EnterLead As New ToolStripMenuItem '' TEMPORARY DEBUG REMOVE LATER
    Friend WithEvents LastID As String
    Public LastIDS As String
    Public Tab As String
    Public cntTotal As Integer = 0
    Public cntConfirmed As Integer = 0
    Public cntUnconfirmed As Integer = 0
    Public cntCandC As Integer = 0
    Public OrigRep1 As String
    Public OrigRep2 As String

    Private Sub Confirming_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        STATIC_VARIABLES.CurrentID = ""
        Me.Dispose()
        Main.Refresh()
    End Sub

    Private Sub Confirming_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing


    End Sub



    Private Sub Confirming_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Tab = "Confirm"
        'Me.WindowState = FormWindowState.Maximized









        Dim c As New ConfirmingData

        c.GetPrimaryLeadSource()
        Me.lvConfirming.Columns(0).Width = 0
        c.PopReps()


        'Dim w As New Open_Windows
        'w.Win_List()





        'Me.MdiParent = Main

        '           Set Expand Characters
        Me.btnExpandConfirming.Text = Chr(187)
        Me.btnExpandSales.Text = Chr(187)

        '           Set Panels
        Me.SplitContainer1.SplitterDistance = 225
        Me.SplitContainer1.IsSplitterFixed = True

        '           Select First Listview Item Both Tabs
        'If Me.lvConfirming.Items.Count > 0 Then
        '    Me.lvConfirming.TopItem.Selected = True
        '    LastID = Me.lvConfirming.SelectedItems(0).Text

        'End If
        'If Me.lvSales.Items.Count > 0 Then
        '    Me.lvSales.TopItem.Selected = True
        'End If
        '           Set Customer History Filter
        Me.TScboCustomerHistory.Text = "All"

        '                       Set up Toolbar Buttons
        Dim x
        x = Me.ilToolStripIcons.Images
        'Confirm
        Me.Confirm.Text = "Confirm Appointment"
        Me.Confirm.Image = x(0)
        Me.Confirmwith.Text = "Select who you spoke with" & vbCr & "to Confirm this Appt."
        'Edit Cutomer
        Me.EditCustomer.Text = "Edit Customer"
        Me.EditCustomer.Image = x(1)
        'Change Status
        Me.ChangeStatus.Text = "Change Status"
        Me.ChangeStatus.Image = x(2)
        'Undo Confirm
        Me.UndoConfirm.Text = "Undo Confirm Appointment"
        Me.UndoConfirm.Image = x(3)
        'Email
        Me.Emailc.Text = "Email Wizard"
        Me.Emailc.Image = x(4)
        Me.Emailu.Text = "Email Wizard"
        Me.Emailu.Image = x(4)
        'Auto dialer
        Me.AutoDial.Text = "Auto Dialer"
        Me.AutoDial.Image = x(5)
        'Sales Result
        Me.Sales.Text = "Edit Sales Appointment"
        Me.Sales.Image = x(6)
        Me.Sales.DropDownItems.Add(Me.SalesResult)
        Me.Sales.DropDownItems.Add(Me.btnCNGApptTime)
        Me.Sales.DropDownItems.Add(Me.SalesSeparator)
        Me.Sales.DropDownItems.Add(Me.Issue)


        Me.Rep1.DropDownStyle = ComboBoxStyle.DropDownList


        Me.SaveChanges.Text = "Save Changes"
        Me.SaveChanges.Font = (New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
        Me.SaveChanges.Enabled = False

        Me.Rep2.DropDownStyle = ComboBoxStyle.DropDownList
        Me.Issue.DropDownItems.Add(Me.Rep1)
        Me.Issue.DropDownItems.Add(Me.Rep2)
        Me.Issue.DropDownItems.Add(Me.SCSeparator)
        Me.Issue.DropDownItems.Add(Me.SaveChanges)
        Me.btnCNGApptTime.Text = "Change Appt. Time"
        Me.tsConfirming.Items.Add(Me.EnterLead)



        ' Send Notes
        Me.SendNotes.Text = "Attach Notes for the" & vbCr & "Issuing Sales Manager"
        Me.SendNotes.Image = x(7)
        ' Reschedule
        Me.Reschedule.Text = "Reschedule Appt. For Another Day and Time"
        Me.Reschedule.Image = x(8)
        ' Kill
        Me.btnKill.Text = "Kill This Appointment"
        Me.btnKill.Image = x(9)
        'Cancelled
        Me.Cancelled.Text = "Log This Appointment as Called & Cancelled"
        Me.Cancelled.Image = x(10)
        'Do Not Call 
        Me.DoNotCall.Text = "Mark as Do Not Call"
        Me.DoNotCall.Image = x(11)
        ' Do Not Mail
        Me.DoNotMail.Text = "Mark as Do Not Mail"
        Me.DoNotMail.Image = x(12)
        ' Do Not Call Or Mail
        Me.DoNotCallOrMail.Text = "Mark as Do Not Call or Mail"
        Me.DoNotCallOrMail.Image = x(13)
        ' Correspondence 
        Me.Correspondence.Text = "Log Correspondence with this Customer"
        Me.Correspondence.Image = x(14)
        'Launch Auto Dialer
        Me.LaunchAutodial.Text = "Launch Progressive Dialer"
        ' main phone
        Me.MainPhone.Text = "Call Main-" & Me.txtHousePhone.Text
        ' Alt Phone 1
        Me.AltPhone1.Text = "Call Alt 1-" & Me.txtaltphone1.Text
        ' ALt Phone 2 
        Me.AltPhone2.Text = "Call Alt 2-" & Me.txtaltphone2.Text
        'Email nested buttons
        Me.EmailConfirmation1.Text = "Send Appt. Confirmation Email to this Customer"
        Me.EmailConfirmationAll.Text = "Send Appt. Confirmation Email to all the Confirmed Customers in My List w/ Valid Email Adresses"
        Me.EmailCustomc.Text = "Email This Customer"
        Me.EmailCustomu.Text = "Email This Customer"
        Me.EmailTemplatec.Text = "Choose Email Template"
        Me.EmailTemplateu.Text = "Choose Email Template"
        Me.CreateNewTemplateu.Text = "Create New Template"
        Me.CreateNewTemplatec.Text = "Create New Template"
        '   Email Setup Unconfirmed
        Dim emu
        emu = Me.Emailu.DropDownItems
        emu.add(Me.EmailCustomu)
        emu.add(Me.EmailTemplateu)
        Me.EmailTemplateu.DropDownItems.Add(Me.TemplateListu)
        Me.EmailTemplateu.DropDownItems.Add(Me.CreateNewTemplateu)

        ' Email Setup Confirmed
        Dim emc = Me.Emailc.DropDownItems
        emc.add(Me.EmailConfirmation1)
        emc.add(Me.EmailConfirmationAll)
        emc.add(Me.EmailCustomc)
        emc.add(Me.EmailTemplatec)

        Me.EmailTemplatec.DropDownItems.Add(Me.TemplateListc)
        '' email template cbo pop
        '' 10-26-15
        GetTemplates()

        Me.EmailTemplatec.DropDownItems.Add(Me.CreateNewTemplatec)

        ' Sales Result
        Me.SalesResult.Text = "Enter Sales Result and Reschedule Appt."
        'Reissue
        Me.Issue.Text = "Issue Appt./Change Rep(s) for this Appt."
        'Rep1
        Me.Rep1.Size = New Size(120, 21)
        Me.Rep1.FlatStyle = FlatStyle.System
        'Rep 2
        Me.Rep2.Size = New Size(120, 21)
        Me.Rep2.FlatStyle = FlatStyle.System
        ' Template List
        Me.TemplateListu.FlatStyle = FlatStyle.System
        Me.TemplateListc.FlatStyle = FlatStyle.System

        'Date Label
        Me.TBLabel.Font = (New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
        Me.TBLabel.Margin = New Padding(0, 1, 115, 2)
        TBLabel.Alignment = ToolStripItemAlignment.Right

        'Date Picker Sales
        Dim z As Integer = Me.tsConfirming.Width
        'MsgBox(Me.dpConfirming.Location.ToString & " |  " & Me.tsConfirming.Width.ToString)



        Me.dpSales.Format = DateTimePickerFormat.Short
        dpSales.Location = New System.Drawing.Point(z - 105, 2)
        dpSales.Size = New System.Drawing.Size(98, 21)
        dpSales.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        dpSales.BringToFront()
        Me.dpSales.Value = Today
        Me.Controls.Add(dpSales)
        Me.dpSales.Visible = False
        'Me.dpSales.alignment = ToolStripItemAlignment.Right


        'Date Picker Confirming
        Me.dpConfirming.Format = DateTimePickerFormat.Short


        dpConfirming.Location = New System.Drawing.Point(z - 105, 2)
        Me.dpConfirming.Margin = New Padding(0, 1, 115, 2)


        'MsgBox(z.ToString)
        'Me.dpConfirming.Location = New System.Drawing.Point(z + 100, 2)
        dpConfirming.Size = New System.Drawing.Size(98, 21)
        dpConfirming.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        dpConfirming.BringToFront()

        Me.dpConfirming.BringToFront()
        Me.dpConfirming.Value = Today.AddDays(1)





        Me.SplitContainer1.Panel2.Controls.Remove(Me.tsAutoDial)
        Me.Controls.Add(Me.tsAutoDial)
        Me.tsAutoDial.Location = New System.Drawing.Point(335, 25)

        c.Populate("Confirm", Me.cboConfirmingPLS.Text.ToString, Me.cboConfirmingSLS.Text.ToString, Me.dpConfirming.Value.ToString, "Populate")
        c.Populate("Dispatch", Me.cboSalesPLS.Text.ToString, Me.cboSalesSLS.Text.ToString, Me.dpSales.Value.ToString, "Populate")
        If Me.lvSales.SelectedItems.Count > 0 Then
            Me.lvSales.TopItem.Selected = True
            LastIDS = Me.lvSales.SelectedItems(0).Text

        End If
        lvConfirming_SelectedIndexChanged(Nothing, Nothing)
        Me.RefreshData.Interval = 5000
        Me.RefreshData.Start()

    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.TabControl1.SelectedIndex = 1 Then
            Tab = "Dispatch"

            Dim c As New ConfirmingData
            c.Populate("Dispatch", Me.cboSalesPLS.Text.ToString, Me.cboSalesSLS.Text.ToString, Me.dpSales.Value.ToString, "Refresh")
            Me.RefreshData.Stop()
            Me.tsConfirming.Items.Clear()
            Me.tsConfirming.Items.Add(Me.EditCustomer)
            Me.tsConfirming.Items.Add(Me.Sales)
            Me.tsConfirming.Items.Add(Me.Emailu)
            Me.tsConfirming.Items.Add(Me.AutoDial)
            Dim ad
            ad = Me.AutoDial.DropDownItems
            ad.add(Me.LaunchAutodial)
            ad.add(Me.Seperator2)
            ad.add(Me.MainPhone)
            Dim alt1 As Integer
            alt1 = Me.txtaltphone1.Text.Length
            Select Case alt1
                Case True
                    Exit Select
                Case False
                    ad.add(Me.AltPhone1)
            End Select
            Dim alt2 As Boolean
            alt2 = Me.txtaltphone2.Text = ""
            Select Case alt2
                Case True
                    Exit Select
            End Select
            Me.TBLabel.Text = "Sales Appointments For"
            Me.tsConfirming.Items.Add(Me.TBLabel)
            Me.dpConfirming.Visible = False
            Me.dpSales.Visible = True
            Me.tsConfirming.SendToBack()
            Dim cs As New ConfirmingData
            If Me.lvSales.SelectedItems.Count > 0 Then
                Me.Text = "Confirming"
                c.PullCustomerINFO("Dispatch", Me.lvSales.SelectedItems(0).Text)
            Else
                c.PullCustomerINFO("Dispatch", "")
            End If
            Me.lvSales_SelectedIndexChanged(Nothing, Nothing)
        Else
            Me.RefreshData.Start()
            Me.RefreshData_Tick(Nothing, Nothing)
            Tab = "Confirm"
            Dim c As New ConfirmingData
            If Me.lvConfirming.SelectedItems.Count > 0 Then
                c.PullCustomerINFO("Confirm", Me.lvConfirming.SelectedItems(0).Text)
            Else
                c.PullCustomerINFO("Confirm", "")
            End If
            Me.dpConfirming.Visible = True
            Me.dpSales.Visible = False
            Me.lvConfirming_SelectedIndexChanged(Nothing, Nothing)
        End If
        'Me.cboConfirmingPLS_LostFocus(Nothing, Nothing)
    End Sub



    Public Sub lvConfirming_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvConfirming.SelectedIndexChanged





        If Me.lvConfirming.SelectedItems.Count > 0 Then
            Dim c As New ConfirmingData

            c.PullCustomerINFO("Confirm", Me.lvConfirming.SelectedItems(0).Text)
            If Me.Tab = "Confirm" Then
                STATIC_VARIABLES.CurrentID = Me.lvConfirming.SelectedItems(0).Text.ToString
                Me.Text = "Confirming"
                Me.Text = Me.Text & " Record ID: " & STATIC_VARIABLES.CurrentID.ToString
            End If
        Else
            Me.Text = "Confirming"
            STATIC_VARIABLES.CurrentID = ""
            'MsgBox(Me.Name)
        End If
        Me.MainPhone.Text = "Call Main Phone - " & Me.txtHousePhone.Text
        If Me.txtaltphone1.Text = "" Then
        Else
            Me.tsbtnHousePhone.DropDownItems.Add(btnAlt1ad)
            Me.btnAlt1ad.Text = "Call Alt Phone 1 - " & Me.txtaltphone1.Text
        End If
        If Me.txtaltphone2.Text = "" Then
        Else
            Me.tsbtnHousePhone.DropDownItems.Add(btnAlt2ad)
            Me.btnAlt2ad.Text = "Call Alt Phone 2 - " & Me.txtaltphone2.Text
        End If
        If Me.lvConfirming.Items.Count > 0 Then
            If Me.lvConfirming.SelectedItems Is Me.lvConfirming.Items(0) Then
                Me.tsbtnPreviousRecord.Enabled = False
            Else
                Me.tsbtnPreviousRecord.Enabled = True
            End If
        End If
        'If
        ' need code to disable next record when last item is selected
        'Me.tsbtnNextRecord.Enabled = False
        'Else 
        'Me.tsbtnNextRecord.Enabled = True
        'End If
        If Me.lvConfirming.SelectedItems.Count = 0 Then
            Me.tsConfirming.Items.Clear()
            Me.tsConfirming.Items.Add(Me.Confirm)


            Me.tsConfirming.Items.Add(Me.EditCustomer)
            Me.tsConfirming.Items.Add(Me.ChangeStatus)
            Dim cs
            cs = Me.ChangeStatus.DropDownItems
            cs.Add(Me.Reschedule)
            cs.add(Me.btnKill)
            cs.add(Me.Seperator)
            cs.add(Me.DoNotCall)
            cs.add(Me.DoNotMail)
            cs.add(Me.DoNotCallOrMail)
            Me.tsConfirming.Items.Add(Me.Emailu)
            Me.tsConfirming.Items.Add(Me.AutoDial)
            Dim ad
            ad = Me.AutoDial.DropDownItems
            ad.add(Me.LaunchAutodial)
            ad.add(Me.Seperator2)
            Me.MainPhone.Text = "Call Main Phone - " & Me.txtHousePhone.Text
            ad.add(Me.MainPhone)
            Dim alt1 As Boolean
            alt1 = Me.txtaltphone1.Text = ""
            Select Case alt1
                Case True
                    Exit Select
                Case False
                    Me.AltPhone1.Text = "Call Alt Phone 1 - " & Me.txtaltphone1.Text
                    ad.add(Me.AltPhone1)
            End Select
            Dim alt2 As Boolean
            alt2 = Me.txtaltphone2.Text = ""
            Select Case alt2
                Case True
                    Exit Select
                Case False
                    Me.AltPhone2.Text = "Call Alt Phone 2 - " & Me.txtaltphone2.Text
                    ad.add(Me.AltPhone2)
            End Select
            Me.TBLabel.Text = "Confirming For"
            Me.tsConfirming.Items.Add(Me.TBLabel)
            Me.Controls.Add(dpConfirming)
            Me.tsConfirming.SendToBack()
            Exit Sub
        End If
        If Me.lvConfirming.SelectedItems(0).Tag = ("Unconfirmed") Or Me.lvConfirming.SelectedItems(0).Tag = "Called and Cancelled" Then
            Me.tsConfirming.Items.Clear()
            Me.tsConfirming.Items.Add(Me.Confirm)

            If Me.txtContact2.Text = " " Then
                Me.Confirmwith.DropDownItems.Clear()
                Me.Confirm.DropDownItems.Add(Me.Confirmwith)
                Me.Confirm.DropDownItems.Add(Me.ConfirmSeparator)
                Me.Confirm.DropDownItems.Add(Me.SendNotes)
                Me.ConfirmwithContact1.Text = Me.txtContact1.Text
                Me.Confirmwith.DropDownItems.Add(Me.ConfirmwithContact1)
                Me.Confirmwithther.Text = "Other..."


                Me.Confirmwith.DropDownItems.Add(Me.Confirmwithther)

            ElseIf Me.txtContact2.Text <> " " Then
                Me.Confirmwith.DropDownItems.Clear()
                Dim s = Split(Me.txtContact1.Text, " ")
                Dim c1 = s(0)
                Dim s2 = Split(Me.txtContact2.Text, " ")
                Dim c2 = s2(0)
                Me.Confirm.DropDownItems.Add(Me.Confirmwith)
                Me.Confirm.DropDownItems.Add(Me.ConfirmSeparator)
                Me.Confirm.DropDownItems.Add(Me.SendNotes)
                Me.ConfirmwithContact1.Text = c1
                Me.Confirmwith.DropDownItems.Add(Me.ConfirmwithContact1)
                Me.ConfirmwithContact2.Text = c2
                Me.Confirmwith.DropDownItems.Add(Me.ConfirmwithContact2)

                Me.ConfirmwithBoth.Text = c1 & " and " & c2
                Me.Confirmwith.DropDownItems.Add(Me.ConfirmwithBoth)
                Me.Confirmwithther.Text = "Other..."
                Me.Confirmwith.DropDownItems.Add(Me.Confirmwithther)

            End If

            Me.tsConfirming.Items.Add(Me.EditCustomer)
            Me.tsConfirming.Items.Add(Me.ChangeStatus)
            Dim cs
            cs = Me.ChangeStatus.DropDownItems
            cs.Add(Me.Reschedule)
            cs.add(Me.btnKill)
            cs.add(Me.Seperator)
            cs.add(Me.DoNotCall)
            cs.add(Me.DoNotMail)
            cs.add(Me.DoNotCallOrMail)
            Me.tsConfirming.Items.Add(Me.Emailu)
            Me.tsConfirming.Items.Add(Me.AutoDial)
            Dim ad
            ad = Me.AutoDial.DropDownItems
            ad.add(Me.LaunchAutodial)
            ad.add(Me.Seperator2)
            Me.MainPhone.Text = "Call Main Phone - " & Me.txtHousePhone.Text
            ad.add(Me.MainPhone)

            Dim alt1 As Boolean
            alt1 = Me.txtaltphone1.Text = ""
            Select Case alt1
                Case True
                    Exit Select
                Case False
                    Me.AltPhone1.Text = "Call Alt Phone 1 - " & Me.txtaltphone1.Text
                    ad.add(Me.AltPhone1)
            End Select
            Dim alt2 As Boolean
            alt2 = Me.txtaltphone2.Text = ""
            Select Case alt2
                Case True
                    Exit Select
                Case False
                    Me.AltPhone2.Text = "Call Alt Phone 2 - " & Me.txtaltphone2.Text
                    ad.add(Me.AltPhone2)
            End Select
            Me.TBLabel.Text = "Confirming For"
            Me.tsConfirming.Items.Add(Me.TBLabel)
            Me.Controls.Add(dpConfirming)
            Me.tsConfirming.SendToBack()
            Exit Sub
        End If
        If Me.lvConfirming.SelectedItems(0).Tag = ("Confirmed") Then
            Me.tsConfirming.Items.Clear()
            Me.tsConfirming.Items.Add(Me.UndoConfirm)
            Me.UndoConfirm.DropDownItems.Add(Me.SendNotes)
            Me.tsConfirming.Items.Add(Me.EditCustomer)
            Me.tsConfirming.Items.Add(Me.Emailc)
            Me.tsConfirming.Items.Add(Me.AutoDial)
            Dim ad
            ad = Me.AutoDial.DropDownItems
            ad.add(Me.LaunchAutodial)
            ad.add(Me.Seperator2)
            ad.add(Me.MainPhone)


            Dim alt1 As Boolean
            alt1 = Me.txtaltphone1.Text = ""
            Select Case alt1
                Case True
                    Exit Select
                Case False
                    Me.AltPhone1.Text = "Call Alt Phone 1 - " & Me.txtaltphone1.Text
                    ad.add(Me.AltPhone1)
            End Select
            Dim alt2 As Boolean
            alt2 = Me.txtaltphone2.Text = ""
            Select Case alt2
                Case True
                    Exit Select
                Case False
                    Me.AltPhone2.Text = "Call Alt Phone 2 - " & Me.txtaltphone2.Text
                    ad.add(Me.AltPhone2)
            End Select
            Me.TBLabel.Text = "Confirming For"
            Me.tsConfirming.Items.Add(Me.TBLabel)
            Me.dpConfirming.Visible = True
            Me.dpSales.Visible = False
            Exit Sub
        End If


    End Sub

    Private Sub btnExpandConfirming_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExpandConfirming.Click
        Select Case Me.btnExpandConfirming.Text
            Case Chr(187)
                Me.SplitContainer1.SplitterDistance = 2500
                Me.SplitContainer1.SplitterWidth = 100
                Me.btnExpandSales.Text = Chr(171)
                Me.btnExpandConfirming.Text = Chr(171)


            Case Chr(171)
                Me.SplitContainer1.SplitterDistance = 225
                Me.SplitContainer1.SplitterWidth = 5
                Me.btnExpandConfirming.Text = Chr(187)
                Me.btnExpandSales.Text = Chr(187)

        End Select

    End Sub

    Private Sub btnExpandSales_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExpandSales.Click
        Me.btnExpandConfirming_Click(Nothing, Nothing)
    End Sub









    Private Sub lblConfimingPLS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblConfimingPLS.Click
        Me.cboConfirmingPLS.Focus()
        Me.cboConfirmingPLS.DroppedDown = True
    End Sub

    Private Sub lblConfirmingSLS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblConfirmingSLS.Click
        Me.cboConfirmingSLS.Focus()
        Me.cboConfirmingSLS.DroppedDown = True
    End Sub







    Private Sub lblSalesPLS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblSalesPLS.Click
        Me.cboSalesPLS.Focus()
        Me.cboSalesPLS.DroppedDown = True
    End Sub



    Private Sub lblSalesSLS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblSalesSLS.Click
        Me.cboSalesSLS.Focus()
        Me.cboSalesSLS.DroppedDown = True
    End Sub

    Private Sub UndoConfirm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UndoConfirm.Click
        If Me.lvConfirming.SelectedItems.Count <= 0 Then
            MsgBox("You must select a record to Undo!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        Dim c As New ConfirmingData
        If Me.dpConfirming.Value < Today Then
            MsgBox("Time has expired to undo this confirmation!", MsgBoxStyle.Exclamation, "Cannot ""Undo"" Confirm Appt.")
            Exit Sub
        End If
        c.Confirm(Me.lvConfirming.SelectedItems(0).Text, sender.text, "", "")
        c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Refresh")
        'Me.lvConfirming_SelectedIndexChanged(Nothing, Nothing)
    End Sub

    Private Sub LaunchAutodial_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LaunchAutodial.Click
        Me.tsAutoDial.Visible = True
        Me.tsAutoDial.BringToFront()
    End Sub

    Private Sub tsbtnCloseDialer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbtnCloseDialer.Click
        Me.tsAutoDial.Visible = False
    End Sub

    Private Sub SendNotes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SendNotes.Click
        If Me.lvConfirming.SelectedItems.Count = 0 Then
            MsgBox("You must select a record to attach note!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        If Me.lvConfirming.SelectedItems.Item(0).Tag = "Confirmed" Then
            SendNotesSM.ID = Me.lvConfirming.SelectedItems(0).Text
            SendNotesSM.ShowInTaskbar = False
            SendNotesSM.ShowDialog()
        Else
            MsgBox("You can only attach a note to a Confirmed Appointment!", MsgBoxStyle.Exclamation, "Appointment Must be Confirmed")
        End If
    End Sub

    Private Sub Reschedule_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Reschedule.Click
        If Me.TabControl1.SelectedIndex = 0 Then
            If Me.lvConfirming.SelectedItems.Count <> 0 Then
                RescheduleAppt.ID = Me.lvConfirming.SelectedItems(0).Text
                RescheduleAppt.frm = "Confirming"
                RescheduleAppt.ShowInTaskbar = False
                RescheduleAppt.ShowDialog()
            Else
                MsgBox("You must select a record to reschedule Appt. Date!", MsgBoxStyle.Exclamation, "No Record Selected")
            End If
        ElseIf Me.TabControl1.SelectedIndex = 1 Then
            If Me.lvSales.SelectedItems.Count <> 0 Then
                RescheduleAppt.ID = Me.lvSales.SelectedItems(0).Text
                RescheduleAppt.frm = "Confirming"
                RescheduleAppt.ShowInTaskbar = False
                RescheduleAppt.ShowDialog()
            Else
                MsgBox("You must select a record to Send Notes!", MsgBoxStyle.Exclamation, "No Record Selected")
            End If
        End If

    End Sub

    Private Sub SalesResult_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SalesResult.Click
        If Me.lvSales.SelectedItems.Count = 0 Then
            MsgBox("You must select a record to enter a result!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        If Me.SalesResult.Text = "Enter Sales Result and Reschedule Appt." Then
            Reissue.ShowInTaskbar = False
            Reissue.ShowDialog()
        Else
            Me.Reschedule_Click(Nothing, Nothing)
        End If

    End Sub

    Private Sub CreateNewTemplateu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CreateNewTemplateu.Click
        EmailTemplate.ShowInTaskbar = False
        EmailTemplate.ShowDialog()
    End Sub

    Private Sub calledcancelled_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles calledcancelled.Click

        Dim s = Split(Me.txtContact1.Text, " ")
        Dim c1 = s(0)
        Dim s2 = Split(Me.txtContact2.Text, " ")
        Dim c2 = s2(0)
        CandCNotes.Contact1 = c1
        CandCNotes.Contact2 = c2
        CandCNotes.OrigApptDate = Me.txtApptDate.Text
        CandCNotes.OrigApptTime = Me.txtApptTime.Text
        CandCNotes.frm = Me

        If Tab = "Confirm" Then
            If Me.lvConfirming.SelectedItems.Count = 0 Then
                MsgBox("You must select an Appt. to Log Cancellation!", MsgBoxStyle.Exclamation, "Cannot Log Cancellation")
            Else
                CandCNotes.ID = Me.lvConfirming.SelectedItems(0).Text
                CandCNotes.ShowInTaskbar = False
                CandCNotes.ShowDialog()
            End If

        ElseIf Tab = "Dispatch" Then
            If Me.lvSales.SelectedItems.Count = 0 Then
                MsgBox("You must select an Appt. to Log Cancellation!", MsgBoxStyle.Exclamation, "Cannot Log Cancellation")
            Else
                CandCNotes.ID = Me.lvSales.SelectedItems(0).Text
                CandCNotes.ShowInTaskbar = False
                CandCNotes.ShowDialog()
            End If
        End If

    End Sub

    Private Sub btnLogCall_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogCall.Click
        If Tab = "Confirm" Then
            If Me.lvConfirming.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Customer" & vbCr & "to Log a Phone Conversation", MsgBoxStyle.Exclamation, "No Customer Selected")
                Exit Sub
            End If
            LogPhoneCall.ID = Me.lvConfirming.SelectedItems(0).Text

        ElseIf Tab = "Dispatch" Then
            If Me.lvSales.SelectedItems.Count = 0 Then
                MsgBox("You must Select a Customer" & vbCr & "to Log a Phone Conversation", MsgBoxStyle.Exclamation, "No Customer Selected")
                Exit Sub
            End If
            LogPhoneCall.ID = Me.lvSales.SelectedItems(0).Text
        End If
        LogPhoneCall.Contact1 = Me.txtContact1.Text
        LogPhoneCall.Contact2 = Me.txtContact2.Text
        LogPhoneCall.frm = Me
        LogPhoneCall.ShowInTaskbar = False
        LogPhoneCall.ShowDialog()
    End Sub

    Private Sub RefreshData_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RefreshData.Tick
        Dim PLS = Me.cboConfirmingPLS.Text
        Dim SLS = Me.cboConfirmingSLS.Text
        Dim ApptDate = Me.dpConfirming.Value
        If PLS = "" Then
            PLS = "%"
        End If
        If SLS = "" Then
            SLS = "%"
        End If
        ApptDate.ToString()
        Dim x = InStr(ApptDate, " ")
        If x <> 0 Then
            ApptDate = Microsoft.VisualBasic.Left(ApptDate, x - 1)
        End If
        ApptDate = ApptDate & " 12:00:00 AM"
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand
        cmdGet = New SqlCommand("dbo.ConfirmingRefreshCheck", cnn)
        Dim r1 As SqlDataReader
        Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)
        Dim param2 As SqlParameter = New SqlParameter("@SLS", SLS)
        Dim param3 As SqlParameter = New SqlParameter("@ApptDate", ApptDate)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cnn.Open()
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r1.Read()
        'MsgBox(r1.Item(0).ToString & " " & r1.Item(1).ToString & " " & r1.Item(2).ToString & " " & r1.Item(3).ToString)
        If r1.Item(0) <> cntTotal Or r1.Item(1) <> cntConfirmed Or r1.Item(2) <> cntUnconfirmed Or r1.Item(3) <> cntCandC Then
            cntTotal = r1.Item(0)
            cntConfirmed = r1.Item(1)
            cntUnconfirmed = r1.Item(2)
            cntCandC = r1.Item(3)
            Dim c As New ConfirmingData
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Refresh")
        End If
        r1.Close()
        cnn.Close()

    End Sub





    Private Sub dpConfirming_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpConfirming.ValueChanged
        Me.RefreshData.Start()
        Dim c As New ConfirmingData
        c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
        If Me.lvConfirming.SelectedItems.Count < 1 Then
            Me.Text = "Confirming"
        End If

    End Sub

    Private Sub TScboCustomerHistory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TScboCustomerHistory.SelectedIndexChanged
        If Tab = "Confirm" Then
            If Me.lvConfirming.SelectedItems.Count <> 0 Then
                Dim c As New CustomerHistory
                c.SetUp(Me, Me.lvConfirming.SelectedItems(0).Text, Me.TScboCustomerHistory)
            End If
        ElseIf Tab = "Dispatch" Then
            If Me.lvSales.SelectedItems.Count <> 0 Then
                Dim c As New CustomerHistory
                c.SetUp(Me, Me.lvSales.SelectedItems(0).Text, Me.TScboCustomerHistory)
            Else
                'Me.lvSales.Items(LastIDS).Selected = True
            End If
        End If
    End Sub

    Private Sub ConfirmwithContact1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ConfirmwithContact1.Click
        If Me.lvConfirming.SelectedItems.Count <= 0 Then
            MsgBox("You must select a record to Confirm!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        Dim c As New ConfirmingData
        Dim c2 As New CustomerHistory
        If Me.dpConfirming.Value < Today Then
            MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Confirm Appt.")
            Exit Sub
        End If
        c.Confirm(Me.lvConfirming.SelectedItems(0).Text, "Confirm Appointment", sender.text, STATIC_VARIABLES.CurrentUser)
        If Me.txtApptDate.Text = Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
            'c.Populate("Dispatch", Me.cboSalesPLS.Text, Me.cboSalesSLS.Text, Me.dpSales.Value.ToString, "Refresh")
            'If Me.lvSales.SelectedItems.Count = 0 Then
            '    Me.lvSales.TopItem.Selected = True
            'End If
        ElseIf Me.txtApptDate.Text <> Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Refresh")
            c2.SetUp(Me, Me.lvConfirming.SelectedItems(0).Text, Me.TScboCustomerHistory)
        End If
    End Sub

    Private Sub Confirm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Confirm.Click

        If Me.lvConfirming.SelectedItems.Count <= 0 Then
            MsgBox("You must select a record to Confirm!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        If Me.dpConfirming.Value < Today Then
            MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Confirm Appt.")
            Exit Sub
        End If


        Me.Confirmwith.Select()
        Me.Confirmwith.ShowDropDown()
        If Me.txtContact2.Text = " " Then
            Me.ConfirmwithContact1.Select()
        End If

    End Sub

    Private Sub ConfirmwithContact2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ConfirmwithContact2.Click
        If Me.lvConfirming.SelectedItems.Count <= 0 Then
            MsgBox("You must select a record to Confirm!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        Dim c As New ConfirmingData
        Dim c2 As New CustomerHistory
        If Me.dpConfirming.Value < Today Then
            MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Confirm Appt.")
            Exit Sub
        End If
        c.Confirm(Me.lvConfirming.SelectedItems(0).Text, "Confirm Appointment", sender.text, STATIC_VARIABLES.CurrentUser)
        If Me.txtApptDate.Text = Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
            'c.Populate("Dispatch", Me.cboSalesPLS.Text, Me.cboSalesSLS.Text, Me.dpSales.Value.ToString, "Refresh")
            'If Me.lvSales.SelectedItems.Count = 0 Then
            '    Me.lvSales.TopItem.Selected = True
            'End If
        ElseIf Me.txtApptDate.Text <> Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Refresh")
            c2.SetUp(Me, Me.lvConfirming.SelectedItems(0).Text, Me.TScboCustomerHistory)
        End If
    End Sub

    Private Sub ConfirmwithBoth_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ConfirmwithBoth.Click
        If Me.lvConfirming.SelectedItems.Count <= 0 Then
            MsgBox("You must select a record to Confirm!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        Dim c As New ConfirmingData
        Dim c2 As New CustomerHistory
        If Me.dpConfirming.Value < Today Then
            MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Confirm Appt.")
            Exit Sub
        End If
        c.Confirm(Me.lvConfirming.SelectedItems(0).Text, "Confirm Appointment", sender.text, STATIC_VARIABLES.CurrentUser)
        If Me.txtApptDate.Text = Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
            'c.Populate("Dispatch", Me.cboSalesPLS.Text, Me.cboSalesSLS.Text, Me.dpSales.Value.ToString, "Refresh")
            'If Me.lvSales.SelectedItems.Count = 0 Then
            '    Me.lvSales.TopItem.Selected = True
            'End If
        ElseIf Me.txtApptDate.Text <> Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Refresh")
            c2.SetUp(Me, Me.lvConfirming.SelectedItems(0).Text, Me.TScboCustomerHistory)
        End If
    End Sub

    Private Sub Confirmwithther_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Confirmwithther.Click
        Dim i As String = InputBox$("Please Enter The Name of the Person" & vbCr & "You Confirmed This Appt. With.", "Enter ""Other"" Contact")
        If i = "" Then
            MsgBox("You must enter a Name!", MsgBoxStyle.Exclamation, "Name Required!")
            Exit Sub
        End If
        Dim c As New ConfirmingData
        Dim c2 As New CustomerHistory
        If Me.dpConfirming.Value < Today Then
            MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Confirm Appt.")
            Exit Sub
        End If
        c.Confirm(Me.lvConfirming.SelectedItems(0).Text, "Confirm Appointment", i, STATIC_VARIABLES.CurrentUser)
        If Me.txtApptDate.Text = Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
            'c.Populate("Dispatch", Me.cboSalesPLS.Text, Me.cboSalesSLS.Text, Me.dpSales.Value.ToString, "Refresh")
            'If Me.lvSales.SelectedItems.Count = 0 Then
            '    Me.lvSales.TopItem.Selected = True
            'End If
        ElseIf Me.txtApptDate.Text <> Me.dpConfirming.Value.ToShortDateString Then
            c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Refresh")
            c2.SetUp(Me, Me.lvConfirming.SelectedItems(0).Text, Me.TScboCustomerHistory)
        End If
    End Sub

    Private Sub btnKill_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnKill.Click
        If Me.lvConfirming.SelectedItems.Count = 0 Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
        Else
            Kill.Contact1 = Me.txtContact1.Text
            Kill.Contact2 = Me.txtContact2.Text
            Kill.frm = "Confirming"
            Kill.ID = Me.lvConfirming.SelectedItems(0).Text
            Kill.ShowInTaskbar = False
            Kill.ShowDialog()

        End If
    End Sub

    Private Sub DoNotCall_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DoNotCall.Click
        If Me.lvConfirming.SelectedItems.Count = 0 Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
        Else
            Dim c As New DoNotCallOrMail
            c.DoNot(Me.lvConfirming.SelectedItems(0).Text, sender.text.ToString)
            Dim c2 As New ConfirmingData
            c2.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")






        End If
    End Sub

    Private Sub DoNotCallOrMail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DoNotCallOrMail.Click
        If Me.lvConfirming.SelectedItems.Count = 0 Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
        Else
            Dim c As New DoNotCallOrMail
            c.DoNot(Me.lvConfirming.SelectedItems(0).Text, sender.text.ToString)
            Dim c2 As New ConfirmingData
            c2.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
        End If
    End Sub

    Private Sub DoNotMail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DoNotMail.Click
        If Me.lvConfirming.SelectedItems.Count = 0 Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
        Else
            Dim c As New DoNotCallOrMail
            c.DoNot(Me.lvConfirming.SelectedItems(0).Text, sender.text.ToString)
            Dim c2 As New ConfirmingData
            c2.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
        End If
    End Sub

    Private Sub lvSales_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvSales.SelectedIndexChanged

        If Me.lvSales.SelectedItems.Count > 0 Then

            If Me.lvSales.SelectedItems(0).Tag = "C&C" Then
                Me.SalesResult.Text = "Reschedule Appt. For Another Day and Time"
                Me.btnCNGApptTime.Enabled = False
                Me.Issue.Enabled = False
            Else
                Me.SalesResult.Text = "Enter Sales Result and Reschedule Appt."
                Me.btnCNGApptTime.Enabled = True
                Me.Issue.Enabled = True
            End If
            Me.Rep1.SelectedItem = Nothing
            Me.Rep2.SelectedItem = Nothing
            Dim f
            If Me.lvSales.SelectedItems(0).SubItems(3).Text.Contains("&") Then
                f = Split(Me.lvSales.SelectedItems(0).SubItems(3).Text, " & ")
                Me.Rep1.SelectedItem = f(0)
                Me.Rep2.SelectedItem = f(1)
                OrigRep1 = f(0)
                OrigRep2 = f(1)
            Else
                Me.Rep1.SelectedItem = Me.lvSales.SelectedItems(0).SubItems(3).Text
                OrigRep1 = Me.lvSales.SelectedItems(0).SubItems(3).Text
                OrigRep2 = ""
                Me.Rep2.SelectedItem = Nothing
            End If








            Dim c As New ConfirmingData
            c.PullCustomerINFO("Dispatch", Me.lvSales.SelectedItems(0).Text)
            If Me.Tab = "Dispatch" Then
                STATIC_VARIABLES.CurrentID = Me.lvSales.SelectedItems(0).Text.ToString
                Me.Text = "Confirming"
                Me.Text = Me.Text & " Record ID: " & STATIC_VARIABLES.CurrentID.ToString
            End If
        Else
            Me.Text = "Confirming"
            STATIC_VARIABLES.CurrentID = ""
        End If

        'STATIC_VARIABLES.CurrentID = Me.lvSales.SelectedItems(0).ToString come back


    End Sub

    Private Sub dpSales_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpSales.LocationChanged
        Dim z As Integer = Me.tsConfirming.Width
        dpSales.Location = New System.Drawing.Point(z - 105, 2)
    End Sub

    Private Sub dpSales_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpSales.ValueChanged
        'Me.lvSales.Items.Clear()
        'Me.lvSales.Groups.Clear()
        Dim c As New ConfirmingData
        c.Populate("Dispatch", Me.cboSalesPLS.Text.ToString, Me.cboSalesSLS.Text.ToString, Me.dpSales.Value.ToString, "Populate")
        If Me.lvSales.SelectedItems.Count < 1 Then
            Me.Text = "Confirming"
        End If
    End Sub

    Private Sub Rep1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Rep1.SelectedIndexChanged
        If Rep1.Text <> "" Then
            Me.Rep2.Enabled = True
        End If
        If Me.Rep1.Text <> OrigRep1 And Me.Rep1.Text <> Me.Rep2.Text Then
            Me.SaveChanges.Enabled = True
            Me.SaveChanges.Select()
        Else
            Me.SaveChanges.Enabled = False
        End If
    End Sub

    Private Sub SaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveChanges.Click
        If Me.lvSales.SelectedItems.Count = 0 Then
            MsgBox("You must select a record to change Sales Rep(s)!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        Dim c As New ConfirmingData
        c.ChangeRep(Me.lvSales.SelectedItems(0).Text, Me.Rep1.Text, Me.Rep2.Text, OrigRep1, OrigRep2)

    End Sub

    Private Sub Sales_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Sales.Click


        If Me.lvSales.SelectedItems.Count = 0 Then
            Me.Rep1.Enabled = False
            Me.Rep2.Enabled = False
        Else
            Me.Rep1.Enabled = True
            Me.Rep2.Enabled = True
        End If

        Me.SaveChanges.Enabled = False
    End Sub

    Private Sub Rep2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Rep2.SelectedIndexChanged

        If Me.Rep2.Text <> OrigRep2 And Me.Rep1.Text <> Me.Rep2.Text Then
            Me.SaveChanges.Enabled = True
            Me.SaveChanges.Select()
        Else
            Me.SaveChanges.Enabled = False
        End If

    End Sub

    Private Sub btnCNGApptTime_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCNGApptTime.Click
        If Me.lvSales.SelectedItems.Count = 0 Then
            MsgBox("You must select a record to change Appt. Time!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        CNGApptTime.ShowInTaskbar = False
        CNGApptTime.ShowDialog()
    End Sub
    Private Sub lnkEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkEmail.Click
        System.Diagnostics.Process.Start("outlook.exe")
    End Sub

    Private Sub EditCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditCustomer.Click
        If Me.TabControl1.SelectedIndex = 0 Then
            If Me.lvConfirming.SelectedItems.Count <> 0 Then
                EditCustomerInfo.ID = Me.lvConfirming.SelectedItems(0).Text
            End If
        Else
            If Me.lvSales.SelectedItems.Count <> 0 Then
                EditCustomerInfo.ID = Me.lvSales.SelectedItems(0).Text
            End If
        End If
        EditCustomerInfo.Show()
    End Sub




    Private Sub tsConfirming_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsConfirming.SizeChanged
        Dim z As Integer = Me.tsConfirming.Width
        dpSales.Location = New System.Drawing.Point(z - 105, 2)

        dpConfirming.Location = New System.Drawing.Point(z - 105, 2)
    End Sub







    'Private Sub btnCNGApptTime_QueryAccessibilityHelp(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryAccessibilityHelpEventArgs) Handles btnCNGApptTime.QueryAccessibilityHelp

    'End Sub

    'Private Sub Confirming_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
    '    Dim z As Integer = Me.tsConfirming.Width
    '    dpSales.Location = New System.Drawing.Point(z - 105, 2)

    '    dpConfirming.Location = New System.Drawing.Point(z - 105, 2)
    'End Sub

    Private Sub Confirming_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'Dim z As Integer = Me.tsConfirming.Width
        'dpSales.Location = New System.Drawing.Point(z - 105, 2)
        'dpConfirming.Location = New System.Drawing.Point(z - 105, 2)
        If Me.lvConfirming.SelectedItems.Count = 0 And Me.TabControl1.SelectedIndex = 0 Then
            Exit Sub
        ElseIf Me.lvSales.SelectedItems.Count = 0 And Me.TabControl1.SelectedIndex = 1 Then
            Exit Sub
        End If
        Dim c As New CustomerHistory
        c.SetUp(Me, Me.lvConfirming.SelectedItems(0).Text, Me.TScboCustomerHistory)
        Me.Refresh()

    End Sub

    Private Sub cboSalesPLS_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSalesPLS.SelectedValueChanged
        'Dim c As New ConfirmingData
        'c.Populate("Dispatch", Me.cboSalesPLS.Text, Me.cboSalesSLS.Text, Me.dpSales.Value.ToString, "Populate")

        If Me.cboSalesPLS.Text <> "" Then
            Me.lblSalesPLS.Text = Me.cboSalesPLS.Text
            Me.lblSalesPLS.ForeColor = Color.Black
            Me.lblSalesPLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.cboSalesSLS.Items.Clear()
            Me.cboSalesSLS_SelectedValueChanged(Nothing, Nothing)
        Else
            Me.lblSalesPLS.ForeColor = Color.Gray
            Me.lblSalesPLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblSalesPLS.Text = "Filter by Primary Lead Source"
            Me.cboSalesSLS.Items.Clear()
            Me.cboSalesSLS_SelectedValueChanged(Nothing, Nothing)

        End If


        If Me.cboSalesPLS.Text <> "" Then
            Dim c As New ConfirmingData
            c.GetSLS(Me.cboSalesPLS.Text)
        End If
        Me.lvSales_SelectedIndexChanged(Nothing, Nothing)
    End Sub
    Private Sub cboConfirmingPLS_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboConfirmingPLS.SelectedValueChanged
        If Me.cboConfirmingPLS.Text <> "" Then
            Me.lblConfimingPLS.Text = Me.cboConfirmingPLS.Text
            Me.lblConfimingPLS.ForeColor = Color.Black
            Me.lblConfimingPLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.cboConfirmingSLS.Items.Clear()
            Me.cboConfirmingSLS_SelectedValueChanged(Nothing, Nothing)
        Else
            Me.lblConfimingPLS.ForeColor = Color.Gray
            Me.lblConfimingPLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblConfimingPLS.Text = "Filter by Primary Lead Source"
            Me.cboConfirmingSLS.Items.Clear()
            Me.cboConfirmingSLS_SelectedValueChanged(Nothing, Nothing)

        End If
        'If Me.lvConfirming.SelectedItems.Count > 0 Then
        '    Dim i = Me.lvConfirming.SelectedItems(0).Text
        '    Me.lvConfirming.Tag = i
        'End If
        Dim c As New ConfirmingData
        ' c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")

        If Me.cboConfirmingPLS.Text <> "" Then
            'Dim c As New ConfirmingData
            c.GetSLS(Me.cboConfirmingPLS.Text)
        End If



    End Sub

    Private Sub cboSalesSLS_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSalesSLS.SelectedValueChanged

        If Me.lvSales.SelectedItems.Count > 0 Then
            Dim i = Me.lvSales.SelectedItems(0).Text
            Me.lvSales.Tag = i
        End If

        Dim c As New ConfirmingData
        c.Populate("Dispatch", Me.cboSalesPLS.Text, Me.cboSalesSLS.Text, Me.dpSales.Value.ToString, "Populate")


        If Me.cboSalesSLS.Text <> "" Then
            Me.lblSalesSLS.Text = Me.cboSalesSLS.Text
            Me.lblSalesSLS.ForeColor = Color.Black
            Me.lblSalesSLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Else
            Me.lblSalesSLS.ForeColor = Color.Gray
            Me.lblSalesSLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblSalesSLS.Text = "Filter by Sec. Lead Source"
        End If
        If Me.lvSales.SelectedItems.Count < 1 Then
            Me.Text = "Confirming"
            STATIC_VARIABLES.CurrentID = ""
        End If

        'If Me.cboSalesPLS.Text <> "" Then
        '    Dim c As New ConfirmingData
        '    c.GetSLS(Me.cboSalesPLS.Text)
        'End If
    End Sub
    Private Sub cboConfirmingSLS_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboConfirmingSLS.SelectedValueChanged




        If Me.cboConfirmingSLS.Text <> "" Then
            Me.lblConfirmingSLS.Text = Me.cboConfirmingSLS.Text
            Me.lblConfirmingSLS.ForeColor = Color.Black
            Me.lblConfirmingSLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Else
            Me.lblConfirmingSLS.ForeColor = Color.Gray
            Me.lblConfirmingSLS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblConfirmingSLS.Text = "Filter by Sec. Lead Source"
        End If





        If Me.lvConfirming.SelectedItems.Count > 0 Then
            Dim i = Me.lvConfirming.SelectedItems(0).Text
            Me.lvConfirming.Tag = i
        End If

        Dim c As New ConfirmingData
        c.Populate(Tab, Me.cboConfirmingPLS.Text, Me.cboConfirmingSLS.Text, Me.dpConfirming.Value.ToString, "Populate")
        If Me.lvConfirming.SelectedItems.Count < 1 Then
            Me.Text = "Confirming"
            STATIC_VARIABLES.CurrentID = ""
        End If
    End Sub

    'Private Sub TabPage1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage1.Click
    '    If Me.lvConfirming.SelectedItems.Count <> 0 Then
    '        Dim c As ConfirmingData
    '        c.PullCustomerINFO(Tab, Me.lvConfirming.SelectedItems(0).Text)
    '    End If
    'End Sub

    'Private Sub TabPage2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Click
    '    If Me.lvSales.SelectedItems.Count <> 0 Then
    '        Dim c As ConfirmingData
    '        c.PullCustomerINFO(Tab, Me.lvSales.SelectedItems(0).Text)
    '    End If
    'End Sub
    Public Sub Create_Tooltip(ByVal w As DateTimePicker, ByVal which As String)

        Dim tt As New ToolTip
        tt.BackColor = Color.White
        tt.IsBalloon = True
        tt.ToolTipTitle = "Date Changed!"
        tt.UseAnimation = True
        tt.UseAnimation = True
        tt.ToolTipIcon = ToolTipIcon.Info
        tt.UseFading = True
        tt.ShowAlways = False
        tt.OwnerDraw = True

        If Me.WindowState = FormWindowState.Maximized Then
            tt.SetToolTip(w, STATIC_VARIABLES.ProgramName & " had to change the date you were " & vbCr & _
                               "working in to link to the Record ID from your " & which)
            'tt.Show("", w)
            tt.Show(STATIC_VARIABLES.ProgramName & " had to change the date you were " & vbCr & _
                   "working in to link to the Record ID from your " & which, _
                   w, -350, -80, 10000)
            Exit Sub
        End If
        tt.Show(STATIC_VARIABLES.ProgramName & " had to change the date you were " & vbCr & _
        "working in to link to the Record ID from your " & which, _
        w, 65, -80, 10000)
    End Sub


    Private Sub rtbSpecialInstructions_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbSpecialInstructions.LostFocus
        Dim ID
        Select Case Me.TabControl1.SelectedIndex
            Case 0
                If Me.lvConfirming.SelectedItems.Count <> 0 Then
                    ID = Me.lvConfirming.SelectedItems(0).Text
                Else
                    Me.rtbSpecialInstructions.Text = ""
                    Exit Sub
                End If
            Case 1
                If Me.lvSales.SelectedItems.Count <> 0 Then
                    ID = Me.lvSales.SelectedItems(0).Text
                Else
                    Me.rtbSpecialInstructions.Text = ""
                    Exit Sub
                End If
            Case Else
                Me.rtbSpecialInstructions.Text = ""
                Exit Sub
        End Select
        Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdUP As SqlCommand = New SqlCommand("Update enterlead set SpecialInstruction = @SI where id = @ID", cnn)
        cmdUP.CommandType = CommandType.Text
        cnn.Open()
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        Dim param2 As SqlParameter = New SqlParameter("@SI", Me.rtbSpecialInstructions.Text)
        cmdUP.Parameters.Add(param1)
        cmdUP.Parameters.Add(param2)
        cmdUP.ExecuteNonQuery()
        cnn.Close()
    End Sub

  

    
    '' email template logic
    '' 10-26-15
    Public Sub GetTemplates()
        Me.TemplateListu.Items.Clear()
        Dim y As New emlTemplateLogic
        Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim depart_ As String = y.GetEmployeeDepartment(name(0), name(1), False)
        Dim g As List(Of emlTemplateLogic.TemplateInfo)
        g = y.GetTemplatesByDepartment(False, depart_)
        Dim a As emlTemplateLogic.TemplateInfo
        For Each a In g
            Me.TemplateListu.Items.Add(a.TemplateName)
        Next

    End Sub

    

    
     
    Private Sub TemplateListu_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TemplateListu.SelectedIndexChanged
        Dim tsCBO As ToolStripComboBox = sender
        Dim ts_txt As String = tsCBO.Text
        If Len(ts_txt) <= 0 Then
            Exit Sub
        ElseIf Len(ts_txt) >= 1 Then
            'MsgBox("Template Name: " & ts_txt, MsgBoxStyle.Information, "DEBUG INFO")
            Dim y As New emlTemplateLogic
            Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim depart_ As String = y.GetEmployeeDepartment(name(0), name(1), False)
            Dim a As emlTemplateLogic.TemplateInfo
            a = y.GetSingleTemplate(ts_txt, False, depart_)
            'MsgBox("Template Name: '" & a.TemplateName & "'" & vbCrLf & "Template Subject: '" & a.Subject & "'" & vbCrLf & "Template Body: '" & a.Body & "'", MsgBoxStyle.Information, "DEBUG INFO")
            Dim b As New emlTemplateLogic
            Dim id As String = STATIC_VARIABLES.CurrentID
            If Len(id) <= 0 Then
                id = b.GetMaxID(False)
            ElseIf Len(id) >= 1 Then
                id = id
            End If
            'Dim scrubbedText As String = b.TestTemplateScrub(id, False, "TEST TEMPLATE", "Administration")
            Dim aa As New convertLeadToStruct
            Dim Lead_ As convertLeadToStruct.EnterLead_Record = aa.ConvertToStructure(id, False)
            frmEmailPreview.LeadToShow = Lead_
            frmEmailPreview.TemplateName = ts_txt
            frmEmailPreview.ShowDialog()
        End If
    End Sub
End Class

