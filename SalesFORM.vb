Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Public Class Sales
#Region "Variables & Dynamic Buttons"
#Region "Criteria Variables"
    Public Rep As String = ""
    Public Product As String = ""
    Public CityState As String = ""
    Public PLS As String = ""
    Public SLS As String = ""
    Public Recovery As Boolean = False
    Public R1 As String = "Demo/No Sale"
    Public R2 As String = "No Demo"
    Public R3 As String = "Reset"
    Public R4 As String = "Not Hit"
    Public R5 As String = "Not Issued"
    Public R6 As String = "Sale"
    Public R7 As String = "Recission Cancel"
    Public R8 As String = "No Results"
    Public panelsize As Integer
#End Region
#Region "Toolbar Buttons"
    '' Summary Tab Button Set 
    Friend WithEvents btnSalesResult As New System.Windows.Forms.ToolStripButton
    Friend WithEvents btnScheduledTasks As New System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents btnMarkTaskAsDone As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnRemoveThisCompletedTask As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnRemoveAllScheduledTask As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnShowAllCompletedTasks As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnSAPreferences As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnEditScheduledTask As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents sepScheduledTasks As New ToolStripSeparator
    Friend WithEvents btnPrintSummary As New ToolStripDropDownButton
    Friend WithEvents btnPrintPerformanceReport As New ToolStripMenuItem
    Friend WithEvents btnPrintNoResultsList As New ToolStripMenuItem
    Friend WithEvents btnPrintScheduledTasks As New ToolStripMenuItem
    Friend WithEvents lblSummary As New ToolStripLabel
    Friend WithEvents dtpSummary As New System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpSummary2 As New System.Windows.Forms.DateTimePicker
    '' Customer List Button Set
    Friend WithEvents btnSalesResult2 As New System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents btnMemorize As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents sepCustomerTools As New ToolStripSeparator
    Friend WithEvents btnMain As New ToolStripMenuItem
    Friend WithEvents btnEMailCustomer As New ToolStripMenuItem
    ''Add more nested Buttons for this 
    ''Add more nested Buttons for this 
    Friend WithEvents btnSetAppt As New ToolStripMenuItem
    Friend WithEvents btnCustomerTools As New System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents btnPrintCustomerList As New ToolStripDropDownButton
    Friend WithEvents btnPrintApptSheet As New ToolStripMenuItem
    Friend WithEvents btnPrintCustomerInfoSheet As New ToolStripMenuItem
    Friend WithEvents btnPrintCurrentList As New ToolStripMenuItem
    Friend WithEvents lblDateRangeCustomerList As New ToolStripLabel
    Friend WithEvents cboDateRangeCustomerList As New ToolStripComboBox
    Friend WithEvents cboDateRangeSummary As New ToolStripComboBox
    Friend WithEvents lblFromCustomerList As New ToolStripLabel
    Friend WithEvents lblFromSummary As New ToolStripLabel
    Friend WithEvents lblToCustomerList As New ToolStripLabel
    Friend WithEvents lblToSummary As New ToolStripLabel
    Friend WithEvents lblIssue As New ToolStripLabel
    Friend WithEvents dtp1CustomerList As New DateTimePicker
    Friend WithEvents dtp2CustomerList As New DateTimePicker
    Friend WithEvents dtpIssueLeads As New DateTimePicker
    Friend WithEvents txt1CustomerList As New TextBox
    Friend WithEvents txt2CustomerList As New TextBox
    Friend WithEvents btnBuildList As New System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents txtSingleRecordInput As New System.Windows.Forms.ToolStripTextBox
    Friend WithEvents btnEditCustomer As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents sepBuildList As New System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnSingleRecord As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnCallCustomer As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnMainPhone As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnAltPhone1 As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnLetter As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnAltPhone2 As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnAssignRep As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnSaveRep As New System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cboRep1 As New System.Windows.Forms.ToolStripComboBox
    Friend WithEvents cboRep2 As New System.Windows.Forms.ToolStripComboBox
    Friend WithEvents sep As New System.Windows.Forms.ToolStripSeparator
    Friend WithEvents sep2 As New System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnMemorizeIssue As New ToolStripButton
    Friend WithEvents btnCCIssue As New ToolStripButton
    Friend WithEvents btnEditCustIssue As New ToolStripButton
    Friend WithEvents btnPrintIssue As New ToolStripDropDownButton
    Friend WithEvents btnEmailIssue As New ToolStripDropDownButton
    Friend WithEvents btnPrintAllIssue As New ToolStripMenuItem
    Friend WithEvents btnPrintNoEmailIssue As New ToolStripMenuItem
    Friend WithEvents btnPrintThisIssue As New ToolStripMenuItem

    Friend WithEvents btnPrintIssuedAppts As New ToolStripMenuItem
    Friend WithEvents btnEmailAllIssue As New ToolStripMenuItem
    Friend WithEvents btnEmailThisIssue As New ToolStripMenuItem
    Friend WithEvents btnEmailIssuedAppts As New ToolStripMenuItem
    Friend WithEvents btnExclude As New ToolStripSplitButton
    Friend WithEvents btnExcludeManage As New ToolStripMenuItem

    '' Issue Leads Button Set 
    '' Attached Files Context Button Set (Item Selected) 

    '' Email template Buttons 
    '' 11-3-2015 AC
    '' 
    



#End Region
#Region "Form Operator Variables"
    Public ID As String = ""
    Public LoadComplete As Boolean = False
    Public LastD1 As String
    Public LastD2 As String
    Public ForceRefresh = False
    Public SingleRecordID As String
    Public TT As ToolTip
    Public SelItem As New ListViewItem
    Public xcordinate As Integer = 0
    Public ycordinate As Integer = 0
    Public IfExists As Boolean = False
#End Region

#End Region

#Region "My Procedures"
    Private Sub ButtonConfig()
        ''Issue Tab Buttons 
        Me.btnCCIssue.Image = Me.ilToolbarButtons.Images(11)
        Me.btnCCIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCCIssue.Name = "btnCCIssue"
        Me.btnCCIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnCCIssue.Text = "Log Appt. as Called and Cancelled"

        Me.btnEditCustIssue.Image = Me.ilToolbarButtons.Images(1)
        Me.btnEditCustIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnEditCustIssue.Name = "btnEditCustIssue"
        Me.btnEditCustIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnEditCustIssue.Text = "Edit Customer"

        Me.btnExclude.Image = Me.ilToolbarButtons.Images(13)
        Me.btnExclude.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnExclude.Name = "btnExclude"
        Me.btnExclude.Size = New System.Drawing.Size(110, 22)
        Me.btnExclude.Text = "Turn Off Exclusions"
        Me.btnExclude.DropDownItems.Add(Me.btnExcludeManage)

        Me.btnExcludeManage.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnExcludeManage.Name = "btnExcludeManage"
        Me.btnExcludeManage.Size = New System.Drawing.Size(110, 22)
        Me.btnExcludeManage.Text = "Manage Rep Info Exclusions"


        Me.btnPrintIssue.Image = Me.ilToolbarButtons.Images(3)
        Me.btnPrintIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrintIssue.Name = "btnPrintIssue"
        Me.btnPrintIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnPrintIssue.Text = "Print"
        Me.btnPrintIssue.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnPrintAllIssue, Me.btnPrintNoEmailIssue, Me.btnPrintThisIssue, Me.btnPrintIssuedAppts})

        Me.btnEmailIssue.Image = Me.ilToolbarButtons.Images(6)
        Me.btnEmailIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnEmailIssue.Name = "btnEmailIssue"
        Me.btnEmailIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnEmailIssue.Text = "Email"
        Me.btnEmailIssue.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnEmailAllIssue, Me.btnEmailThisIssue})

        'Me.btnPrintAllIssue.Image = Me.ilToolbarButtons.Images(1)
        Me.btnPrintAllIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrintAllIssue.Name = "btnPrintAllIssue"
        Me.btnPrintAllIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnPrintAllIssue.Text = "Print All Leads"

        Me.btnPrintNoEmailIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrintNoEmailIssue.Name = "btnPrintNoEmailIssue"
        Me.btnPrintNoEmailIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnPrintNoEmailIssue.Text = "Print Leads for Reps that do not recieve emails"

        Me.btnPrintThisIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrintThisIssue.Name = "btnPrintThisIssue"
        Me.btnPrintThisIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnPrintThisIssue.Text = "Print Selected Lead"

        Me.btnPrintIssuedAppts.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrintIssuedAppts.Name = "btnPrintIssuedAppts"
        Me.btnPrintIssuedAppts.Size = New System.Drawing.Size(110, 22)
        Me.btnPrintIssuedAppts.Text = "Print Issued Leads List" ''

        Me.btnEmailAllIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnEmailAllIssue.Name = "btnEmailAllIssue"
        Me.btnEmailAllIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnEmailAllIssue.Text = "Email Leads to all Reps that recieve email"

        Me.btnEmailThisIssue.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnEmailThisIssue.Name = "btnEmailThisIssue"
        Me.btnEmailThisIssue.Size = New System.Drawing.Size(110, 22)
        Me.btnEmailThisIssue.Text = "Email Selected Lead"

        Me.btnEmailIssuedAppts.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnEmailIssuedAppts.Name = "btnEmailIssuedAppts"
        Me.btnEmailIssuedAppts.Size = New System.Drawing.Size(110, 22)
        Me.btnEmailIssuedAppts.Text = "Email Issued Leads List to Sales Manager"


        '' Summary Tab Buttons
        '' btnSalesResult
        Me.btnSalesResult.Image = Me.ilToolbarButtons.Images(0)
        Me.btnSalesResult.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSalesResult.Name = "btnSalesResult"
        Me.btnSalesResult.Size = New System.Drawing.Size(110, 22)
        Me.btnSalesResult.Text = "Enter sales result"
        'btnScheduledTasks
        '
        Me.btnScheduledTasks.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnMarkTaskAsDone, Me.btnEditScheduledTask, Me.btnRemoveThisCompletedTask, Me.btnRemoveAllScheduledTask, Me.btnShowAllCompletedTasks, Me.sepScheduledTasks, Me.btnSAPreferences})
        Me.btnScheduledTasks.Image = Me.ilToolbarButtons.Images(2)
        Me.btnScheduledTasks.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnScheduledTasks.Name = "btnScheduledTasks"
        Me.btnScheduledTasks.Size = New System.Drawing.Size(131, 22)
        Me.btnScheduledTasks.Text = "Edit Scheduled Task"
        'btnMarkTaskAsDone
        '
        Me.btnMarkTaskAsDone.Name = "btnMarkTaskAsDone"
        Me.btnMarkTaskAsDone.Size = New System.Drawing.Size(209, 22)
        Me.btnMarkTaskAsDone.Text = "Mark Task as Done"
        'btnEditScheduledTask
        '
        Me.btnEditScheduledTask.Name = "btnEditScheduledTask"
        Me.btnEditScheduledTask.Size = New System.Drawing.Size(209, 22)
        Me.btnEditScheduledTask.Text = "Edit Scheduled Task"
        '
        'btnRemoveThisCompletedTask
        '
        Me.btnRemoveThisCompletedTask.Name = "btnRemoveThisCompletedTask"
        Me.btnRemoveThisCompletedTask.Size = New System.Drawing.Size(209, 22)
        Me.btnRemoveThisCompletedTask.Text = "Hide This Completed Task"
        Me.btnRemoveThisCompletedTask.Visible = False
        '
        'btnRemoveAllScheduledTask
        '
        Me.btnRemoveAllScheduledTask.Name = "btnRemoveAllScheduledTask"
        Me.btnRemoveAllScheduledTask.Size = New System.Drawing.Size(209, 22)
        Me.btnRemoveAllScheduledTask.Text = "Hide All Completed Tasks"
        '
        'btnShowAllCompletedTasks
        '
        Me.btnShowAllCompletedTasks.Checked = True
        Me.btnShowAllCompletedTasks.CheckState = System.Windows.Forms.CheckState.Checked
        Me.btnShowAllCompletedTasks.Name = "btnShowAllCompletedTasks"
        Me.btnShowAllCompletedTasks.Size = New System.Drawing.Size(209, 22)
        Me.btnShowAllCompletedTasks.Text = "Show All Completed Tasks"
        'btnSAPreferences
        '
        Me.btnSAPreferences.Name = "btnSAPreferences"
        Me.btnSAPreferences.Size = New System.Drawing.Size(209, 22)
        Me.btnSAPreferences.Text = "Preferences"
        'btnPrintSummary
        '
        Me.btnPrintSummary.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnPrintPerformanceReport, Me.btnPrintPerformanceReport, Me.btnPrintScheduledTasks})
        Me.btnPrintSummary.Image = Me.ilToolbarButtons.Images(3)
        Me.btnPrintSummary.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrintSummary.Name = "btnPrintSummary"
        Me.btnPrintSummary.Size = New System.Drawing.Size(58, 22)
        Me.btnPrintSummary.Text = "Print"
        'btnPrintPerformanceReport
        Me.btnPrintPerformanceReport.Name = "btnPrintPerformanceReport"
        Me.btnPrintPerformanceReport.Size = New System.Drawing.Size(233, 22)
        Me.btnPrintPerformanceReport.Text = "Print Daily Performance Report"
        'btnPrintPerformanceReport
        '
        Me.btnPrintPerformanceReport.Name = "btnPrintPerformanceReport"
        Me.btnPrintPerformanceReport.Size = New System.Drawing.Size(233, 22)
        Me.btnPrintPerformanceReport.Text = "Print ""No Sales Results"" List"
        'btnPrintScheduledTasks
        '
        Me.btnPrintScheduledTasks.Name = "btnPrintScheduledTasks"
        Me.btnPrintScheduledTasks.Size = New System.Drawing.Size(233, 22)
        Me.btnPrintScheduledTasks.Text = "Print ""Scheduled Task"" List"
        'lblSummary
        '
        Me.lblSummary.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblSummary.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        'Me.lblSummary.Margin = New System.Windows.Forms.Padding(0, 1, 115, 2)
        Me.lblSummary.Name = "tslblPerformance"
        Me.lblSummary.Size = New System.Drawing.Size(215, 22)
        Me.lblSummary.Text = "Performance Report Dates:"
        'dtpSummary
        '
        Me.dtpSummary.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpSummary.CalendarFont = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpSummary.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpSummary.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpSummary.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 108, 2)
        Me.dtpSummary.Name = "dtpSummary"
        Me.dtpSummary.Size = New System.Drawing.Size(98, 21)
        Me.dtpSummary.TabIndex = 443
        'Dim ts As New TimeSpan(1, 0, 0, 0)
        'Me.dtpSummary.Value = Me.dtpSummary.Value.Subtract(ts)

        Me.dtpSummary2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpSummary2.CalendarFont = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpSummary2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpSummary2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpSummary2.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 247, 2)
        Me.dtpSummary2.Name = "dtpSummary2"
        Me.dtpSummary2.Size = New System.Drawing.Size(98, 21)
        Me.dtpSummary2.TabIndex = 444
        'Dim ts As New TimeSpan(1, 0, 0, 0)
        'Me.dtpSummary2.Value = Me.dtpSummary2.Value.Subtract(ts)
        Me.cboDateRangeSummary.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.cboDateRangeSummary.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDateRangeSummary.FlatStyle = System.Windows.Forms.FlatStyle.Standard
        Me.cboDateRangeSummary.Items.AddRange(New Object() {"Today", "Yesterday", "This Week", "This Week - to date", "This Month", "This Month - to date", "This Year", "This Year - to date", "Last Week", "Last Week - to date", "Last Month", "Last Month - to date", "Last Year", "Last Year - to date", "Custom"})
        Me.cboDateRangeSummary.Margin = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.cboDateRangeSummary.Name = "cboDateRangeSummary"
        Me.cboDateRangeSummary.Size = New System.Drawing.Size(123, 25)

        Me.lblFromSummary.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblFromSummary.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromSummary.Margin = New System.Windows.Forms.Padding(0, 1, 120, 2)
        Me.lblFromSummary.Name = "lblFromSummary"
        Me.lblFromSummary.Size = New System.Drawing.Size(39, 22)
        Me.lblFromSummary.Text = "From"

        Me.lblToSummary.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblToSummary.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToSummary.Margin = New System.Windows.Forms.Padding(0, 1, 115, 2)
        Me.lblToSummary.Name = "lblToSummary"
        Me.lblToSummary.Size = New System.Drawing.Size(23, 22)
        Me.lblToSummary.Text = "To"


        '' Customer List Tab Buttons
        'btnSalesResult2
        '
        Me.btnSalesResult2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnMemorize})
        Me.btnSalesResult2.Image = Me.ilToolbarButtons.Images(0)
        Me.btnSalesResult2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSalesResult2.Name = "btnSalesResult2"
        Me.btnSalesResult2.Size = New System.Drawing.Size(122, 22)
        Me.btnSalesResult2.Text = "Enter sales result"
        '' Email template wizard button adds
        '' 11-3-15 AC
        ''
        'cboTemplateLists   '' combo box with a list of templates per department to quick apply
        'btnEmailWizard   '' lead in button for all other choices
        'btnEmailTemplateAll   '' will take a list of recID's , apply template, and bulk mail
        'btnEmailTemplateOne 
 




        '
        'btnMemorize
        '
        Me.btnMemorize.Image = Me.ilToolbarButtons.Images(4)
        Me.btnMemorize.Name = "btnMemorize"
        Me.btnMemorize.Size = New System.Drawing.Size(182, 22)
        Me.btnMemorize.Text = "Memorize This Record"
        'btnCustomerTools
        '
        Me.btnCustomerTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnEditCustomer, Me.sepCustomerTools, Me.btnCallCustomer, Me.btnEMailCustomer, Me.btnLetter, Me.sep, Me.btnSetAppt, Me.btnAssignRep})
        Me.btnCustomerTools.Image = Me.ilToolbarButtons.Images(1)
        Me.btnCustomerTools.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCustomerTools.Name = "btnCustomerTools"
        Me.btnCustomerTools.Size = New System.Drawing.Size(110, 22)
        Me.btnCustomerTools.Text = "Customer Tools"
        '
        'btnEditCustomer
        '

        Me.btnEditCustomer.Image = Me.ilToolbarButtons.Images(1)
        Me.btnEditCustomer.Name = "btnEditCustomer"
        Me.btnEditCustomer.Size = New System.Drawing.Size(173, 22)
        Me.btnEditCustomer.Text = "Edit Customer"
        '
        'sepCustomerTools
        '
        Me.sepCustomerTools.Name = "Me.sepCustomerTools"
        Me.sepCustomerTools.Size = New System.Drawing.Size(170, 6)
        '
        'btnCallCustomer
        '
        Me.btnCallCustomer.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnMain, Me.btnAltPhone1, Me.btnAltPhone2})
        Me.btnCallCustomer.Image = Me.ilToolbarButtons.Images(5)
        Me.btnCallCustomer.Name = "btnCallCustomer"
        Me.btnCallCustomer.Size = New System.Drawing.Size(173, 22)
        Me.btnCallCustomer.Text = "Call This Customer"
        '
        'btnMainPhone
        '
        Me.btnMainPhone.Name = "btnMainPhone"
        Me.btnMainPhone.Size = New System.Drawing.Size(248, 22)
        Me.btnMainPhone.Text = ""
        '
        'btnAltPhone1
        '
        Me.btnAltPhone1.Name = "btnAltPhone1"
        Me.btnAltPhone1.Size = New System.Drawing.Size(248, 22)
        Me.btnAltPhone1.Text = ""
        '
        'btnAltPhone2
        '
        Me.btnAltPhone2.Name = "btnAltPhone2"
        Me.btnAltPhone2.Size = New System.Drawing.Size(248, 22)
        Me.btnAltPhone2.Text = ""
        '
        'btnEmailCustomer
        '
        'Me.btnEmailCustomer.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {    ''add range later           })
        Me.btnEMailCustomer.Image = Me.ilToolbarButtons.Images(6)
        Me.btnEMailCustomer.Name = "btnEmailCustomer"
        Me.btnEMailCustomer.Size = New System.Drawing.Size(173, 22)
        Me.btnEMailCustomer.Text = "Email Customer"

        '
        'btnLetter
        '
        'Me.btnLetter.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {  ''Add Range Later         })
        Me.btnLetter.Image = Me.ilToolbarButtons.Images(7)
        Me.btnLetter.Name = "btnLetter"
        Me.btnLetter.Size = New System.Drawing.Size(173, 22)
        Me.btnLetter.Text = "Send a Letter"
        'btnSetAppt
        '
        Me.btnSetAppt.Image = Me.ilToolbarButtons.Images(8)
        Me.btnSetAppt.Name = "btnSetAppt"
        Me.btnSetAppt.Size = New System.Drawing.Size(173, 22)
        Me.btnSetAppt.Text = "Set Appointment"
        'btnAssignRep
        '
        Me.btnAssignRep.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cboRep1, Me.cboRep2, Me.sep2, Me.btnSaveRep})
        Me.btnAssignRep.Name = "btnAssignRep"
        Me.btnAssignRep.Size = New System.Drawing.Size(173, 22)
        Me.btnAssignRep.Text = "Assign/Change Sales Rep(s)"
        'btnSaveRep
        '
        Me.btnSaveRep.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveRep.Name = "btnAssignRep"
        Me.btnSaveRep.Size = New System.Drawing.Size(173, 22)
        Me.btnSaveRep.Text = "Save Changes"
        Me.btnSaveRep.TextAlign = ContentAlignment.MiddleCenter

        'cborep1 and cborep2
        Me.cboRep1.DropDownStyle = ComboBoxStyle.DropDownList
        Me.cboRep1.FlatStyle = FlatStyle.Standard
        Me.cboRep2.DropDownStyle = ComboBoxStyle.DropDownList
        Me.cboRep2.FlatStyle = FlatStyle.Standard
        Me.cboRep1.Size = New System.Drawing.Size(150, Me.cboRep1.Size.Height)
        Me.cboRep2.Size = New System.Drawing.Size(150, Me.cboRep1.Size.Height)
        'btnBuildList
        '
        Me.btnBuildList.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.txtSingleRecordInput, Me.sepBuildList, Me.btnSingleRecord})
        Me.btnBuildList.Image = Me.ilToolbarButtons.Images(9)
        Me.btnBuildList.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnBuildList.Name = "btnBuildList"
        Me.btnBuildList.Size = New System.Drawing.Size(118, 22)
        Me.btnBuildList.Text = "Create Custom List"
        'txtSingleRecordInput
        '

        Me.txtSingleRecordInput.Name = "txtSingleRecordInput"
        Me.txtSingleRecordInput.Size = New System.Drawing.Size(118, 22)
        Me.txtSingleRecordInput.Text = "[Enter Record ID]"
        Me.txtSingleRecordInput.BorderStyle = BorderStyle.FixedSingle
        'btnSingleRecord
        '
        Me.btnSingleRecord.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSingleRecord.TextAlign = ContentAlignment.MiddleCenter
        Me.btnSingleRecord.Name = "btnSingleRecord"
        Me.btnSingleRecord.Size = New System.Drawing.Size(118, 22)
        Me.btnSingleRecord.Text = "Get Record"
        Me.btnSingleRecord.Enabled = False
        'btnPrintCustomerList
        '
        Me.btnPrintCustomerList.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnPrintApptSheet, Me.btnPrintCustomerInfoSheet, Me.btnPrintCurrentList})
        Me.btnPrintCustomerList.Image = Me.ilToolbarButtons.Images(3)
        Me.btnPrintCustomerList.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnPrintCustomerList.Name = "btnPrintCustomerList"
        Me.btnPrintCustomerList.Size = New System.Drawing.Size(58, 22)
        Me.btnPrintCustomerList.Text = "Print"
        'btnPrintApptSheet
        Me.btnPrintApptSheet.Name = "btnPrintApptSheet"
        Me.btnPrintApptSheet.Size = New System.Drawing.Size(233, 22)
        Me.btnPrintApptSheet.Text = "Print Appointment Sheet"
        'btnPrintCustomerInfoSheet
        Me.btnPrintCustomerInfoSheet.Name = "btnPrintCustomerInfoSheet"
        Me.btnPrintCustomerInfoSheet.Size = New System.Drawing.Size(233, 22)
        Me.btnPrintCustomerInfoSheet.Text = "Print Customer Information Sheet"
        '' btnPrintCurrentList
        Me.btnPrintCurrentList.Name = "btnPrintCurrentList"
        Me.btnPrintCurrentList.Size = New System.Drawing.Size(233, 22)
        Me.btnPrintCurrentList.Text = "Print Contact List From The Current Customer List"
        'lblDateRangeCustomerList
        '
        Me.lblDateRangeCustomerList.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblDateRangeCustomerList.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateRangeCustomerList.Name = "lblDateRangeCustomerList"
        Me.lblDateRangeCustomerList.Size = New System.Drawing.Size(86, 22)
        Me.lblDateRangeCustomerList.Text = "Appt. Dates"
        'cboDateRangeCustomerList
        '
        Me.cboDateRangeCustomerList.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.cboDateRangeCustomerList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDateRangeCustomerList.FlatStyle = System.Windows.Forms.FlatStyle.Standard
        Me.cboDateRangeCustomerList.Items.AddRange(New Object() {"All", "Today", "Yesterday", "This Week", "This Week - to date", "This Month", "This Month - to date", "This Year", "This Year - to date", "Next Week", "Next Month", "Last Week", "Last Week - to date", "Last Month", "Last Month - to date", "Last Year", "Last Year - to date", "Custom"})
        Me.cboDateRangeCustomerList.Margin = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.cboDateRangeCustomerList.Name = "cboDateRangeCustomerList"
        Me.cboDateRangeCustomerList.Size = New System.Drawing.Size(123, 25)
        'lblToCustomerList
        '
        Me.lblToCustomerList.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblToCustomerList.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToCustomerList.Margin = New System.Windows.Forms.Padding(0, 1, 115, 2)
        Me.lblToCustomerList.Name = "lblToCustomerList"
        Me.lblToCustomerList.Size = New System.Drawing.Size(23, 22)
        Me.lblToCustomerList.Text = "To"
        '
        'lblFromCustomerList
        '
        Me.lblFromCustomerList.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblFromCustomerList.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromCustomerList.Margin = New System.Windows.Forms.Padding(0, 1, 120, 2)
        Me.lblFromCustomerList.Name = "lblFromCustomerList"
        Me.lblFromCustomerList.Size = New System.Drawing.Size(39, 22)
        Me.lblFromCustomerList.Text = "From"
        'dtp1CustomerList
        '
        Me.dtp1CustomerList.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtp1CustomerList.CalendarFont = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp1CustomerList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp1CustomerList.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp1CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 247, 2) '745
        Me.dtp1CustomerList.Name = "dtp1CustomerList"
        Me.dtp1CustomerList.Size = New System.Drawing.Size(98, 21)
        Me.dtp1CustomerList.TabIndex = 442
        Me.dtp1CustomerList.Visible = False
        'dtp2CustomerList
        '
        Me.dtp2CustomerList.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtp2CustomerList.CalendarFont = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp2CustomerList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp2CustomerList.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp2CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 108, 2) '882
        Me.dtp2CustomerList.Name = "dtp2CustomerList"
        Me.dtp2CustomerList.Size = New System.Drawing.Size(98, 21)
        Me.dtp2CustomerList.TabIndex = 443
        Me.dtp2CustomerList.Visible = False
        'txt1CustomerList
        '
        Me.txt1CustomerList.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt1CustomerList.BackColor = System.Drawing.Color.White
        Me.txt1CustomerList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt1CustomerList.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt1CustomerList.ForeColor = System.Drawing.Color.Black
        Me.txt1CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 246, 4) '747
        Me.txt1CustomerList.Name = "txt1CustomerList"
        Me.txt1CustomerList.ReadOnly = True
        Me.txt1CustomerList.Size = New System.Drawing.Size(60, 16)
        Me.txt1CustomerList.TabIndex = 179
        Me.txt1CustomerList.TabStop = False
        Me.txt1CustomerList.Visible = False
        'txt2CustomerList
        '
        Me.txt2CustomerList.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt2CustomerList.BackColor = System.Drawing.Color.White
        Me.txt2CustomerList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt2CustomerList.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt2CustomerList.ForeColor = System.Drawing.Color.Black
        Me.txt2CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 107, 4) '885
        Me.txt2CustomerList.Name = "txt2CustomerList"
        Me.txt2CustomerList.ReadOnly = True
        Me.txt2CustomerList.Size = New System.Drawing.Size(60, 16)
        Me.txt2CustomerList.TabIndex = 180
        Me.txt2CustomerList.TabStop = False
        Me.txt2CustomerList.Visible = False

        ''IssueLeads Tab
        Me.dtpIssueLeads.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpIssueLeads.CalendarFont = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpIssueLeads.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpIssueLeads.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpIssueLeads.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 108, 2)
        Me.dtpIssueLeads.Name = "dtpIssueLeads"
        Me.dtpIssueLeads.Size = New System.Drawing.Size(98, 21)
        Me.dtpIssueLeads.TabIndex = 442
        Me.dtpIssueLeads.Visible = False
        Me.dtpIssueLeads.Value = Today

        Me.lblIssue.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblIssue.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIssue.Name = "lblIssue"
        Me.lblIssue.Size = New System.Drawing.Size(296, 22)
        Me.lblIssue.Text = "Issue Appointments For"
        Me.lblIssue.Padding = New Padding(0, 0, 228, 0)



        '' Add controls overlaying toolbar
        Me.Controls.Add(Me.dtpSummary)
        Me.Controls.Add(Me.dtpSummary2)
        Me.Controls.Add(Me.dtp1CustomerList)
        Me.Controls.Add(Me.dtp2CustomerList)
        Me.Controls.Add(Me.txt1CustomerList)
        Me.Controls.Add(Me.txt2CustomerList)
        Me.Controls.Add(Me.dtpIssueLeads)

        '' add more datepickers and textboxes as needed for additional tabs
    End Sub
    Public Sub ToolbarConfig(ByVal Toolbar As Integer)
        Me.tsSalesDepartment.Items.Clear()
        Select Case Toolbar
            Case Is = 1  ''Summary Tab
                Me.tsSalesDepartment.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnSalesResult, Me.btnScheduledTasks, Me.btnPrintSummary, Me.lblToSummary, Me.lblFromSummary, Me.cboDateRangeSummary, Me.lblSummary})
                Me.dtpSummary.Visible = True
                Me.dtpSummary.BringToFront()
                Me.dtpSummary2.Visible = True
                Me.dtpSummary2.BringToFront()
                Me.dtp1CustomerList.Visible = False
                Me.dtp2CustomerList.Visible = False
                Me.txt1CustomerList.Visible = False
                Me.txt2CustomerList.Visible = False
                Me.dtpIssueLeads.Visible = False

            Case Is = 2  '' Customer List Tab , Sales List Tab 
                Me.tsSalesDepartment.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnSalesResult2, Me.btnCustomerTools, Me.btnBuildList, Me.btnPrintCustomerList, Me.lblToCustomerList, Me.lblFromCustomerList, Me.cboDateRangeCustomerList, Me.lblDateRangeCustomerList})
                Me.btnMemorize.Text = "Memorize This Record"
                Me.btnMemorize.Image = Me.ilToolbarButtons.Images(4)
                Me.dtpSummary.Visible = False
                Me.dtpSummary2.Visible = False
                Me.dtpIssueLeads.Visible = False
                Me.dtp1CustomerList.Visible = True
                Me.dtp1CustomerList.BringToFront()
                Me.dtp2CustomerList.Visible = True
                Me.dtp2CustomerList.BringToFront()
                Me.txt1CustomerList.BringToFront()
                Me.txt2CustomerList.BringToFront()
                If Me.txt1CustomerList.Text = "" Then
                    Me.txt1CustomerList.Visible = True
                End If
                If Me.txt2CustomerList.Text = "" Then
                    Me.txt2CustomerList.Visible = True
                End If
                Me.ShowNotesToolStripMenuItem.Visible = False
                Me.ToolStripSeparator1.Visible = False
                Me.RemoveThisApptToolStripMenuItem.Visible = False
                Me.MemorizeThisApptToolStripMenuItem.Visible = True

                '' will probably have to revisit this 

                ''Need code to deselect any tasks in the task manager
            Case Is = 3 '' Customer List Tab , Memorized Appts Tab

                Me.tsSalesDepartment.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnSalesResult2, Me.btnCustomerTools, Me.btnPrintCustomerList})
                
                Me.btnMemorize.Text = "Remove This Record"
                Me.btnMemorize.Image = Me.ilToolbarButtons.Images(10)
                Me.txt1CustomerList.Visible = False
                Me.txt2CustomerList.Visible = False
                Me.dtp1CustomerList.Visible = False
                Me.dtp2CustomerList.Visible = False
                Me.dtpSummary.Visible = False
                Me.dtpSummary2.Visible = False
                Me.dtpIssueLeads.Visible = False
                Me.ShowNotesToolStripMenuItem.Visible = True
                Me.ToolStripSeparator1.Visible = True
                Me.RemoveThisApptToolStripMenuItem.Visible = True
                Me.MemorizeThisApptToolStripMenuItem.Visible = False
                '' add email associated buttons here for 
                '' email templates and bulk mailing per list
                '' 11-3-2015 AC
                ''

                '' cboTemplateList 

                '' GetTemplates() -> list of templates
                



            Case Is = 4 '' Issue Leads Tab 
                Me.tsSalesDepartment.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnCCIssue, Me.btnEditCustIssue, Me.btnPrintIssue, Me.btnEmailIssue, Me.btnExclude, Me.lblIssue})
                Me.txt1CustomerList.Visible = False
                Me.txt2CustomerList.Visible = False
                Me.dtp1CustomerList.Visible = False
                Me.dtp2CustomerList.Visible = False
                Me.dtpSummary.Visible = False
                Me.dtpSummary2.Visible = False
                Me.dtpIssueLeads.Visible = True
                Me.dtpIssueLeads.BringToFront()

            Case Is = 5 ''References Tab 
                Me.txt1CustomerList.Visible = False
                Me.txt2CustomerList.Visible = False
                Me.dtp1CustomerList.Visible = False
                Me.dtp2CustomerList.Visible = False
                Me.dtpSummary.Visible = False
                Me.dtpSummary2.Visible = False
                Me.dtpIssueLeads.Visible = False

            Case Is = 6  '' Reports Tab 
                Me.txt1CustomerList.Visible = False
                Me.txt2CustomerList.Visible = False
                Me.dtp1CustomerList.Visible = False
                Me.dtp2CustomerList.Visible = False
                Me.dtpSummary.Visible = False
                Me.dtpSummary2.Visible = False
                Me.dtpIssueLeads.Visible = False

        End Select
    End Sub
    Public Sub PopulateNoResults()
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.NoResultsSummary", Cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        Cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        Dim x As String
        If Me.lvnoresults.SelectedItems.Count <> 0 Then
            x = Me.lvnoresults.SelectedItems(0).SubItems(1).Text
        End If
        Me.lvnoresults.Items.Clear()

        While r1.Read
            Dim lv As New ListViewItem
            Dim d
            Dim rep1 As String
            Dim rep2 As String
            Try
                rep1 = r1.Item(2).ToString
            Catch ex As Exception
                rep1 = ""
            End Try
            Try
                rep2 = r1.Item(3).ToString
            Catch ex As Exception
                rep2 = ""
            End Try

            d = Split(r1.Item(0), " ", 2)
            lv.Text = d(0)
            lv.Tag = r1.Item(1)
            lv.SubItems.Add(r1.Item(1).ToString)
            If rep2 <> "" Then
                lv.SubItems.Add(rep1 & " & " & rep2)
            Else
                lv.SubItems.Add(rep1)
            End If

            Me.lvnoresults.Items.Add(lv)
            If lv.SubItems(1).Text = x Then
                lv.Selected = True
                lv.EnsureVisible()
            End If
        End While
        r1.Close()
        Cnn.Close()
    End Sub
    Public Sub PullInfo(ByVal ID As String)
        If Me.ID = ID And Me.ForceRefresh = False Then
            Exit Sub
        End If
        Me.ID = ID
        STATIC_VARIABLES.CurrentID = Me.ID
        Me.ForceRefresh = False
        Me.btnCallCustomer.DropDownItems.Item(0).Visible = False
        Me.btnCallCustomer.DropDownItems.Item(1).Visible = False
        Me.btnCallCustomer.DropDownItems.Item(2).Visible = False
        If ID = "" Then
            Me.txtContact1.Text = ""
            Me.txtContact2.Text = ""
            Me.txtAddress.Text = ""
            Me.txtWorkHours.Text = ""
            Me.txtHousePhone.Text = ""
            Me.txtaltphone1.Text = ""
            Me.txtaltphone2.Text = ""
            Me.txtAlt1Type.Text = ""
            Me.txtAlt2Type.Text = ""
            Me.lnkEmail.Text = ""
            Me.txtApptDate.Text = ""
            Me.txtApptDay.Text = ""
            Me.txtApptTime.Text = ""
            Me.txtProducts.Text = ""
            Me.txtColor.Text = ""
            Me.txtQty.Text = ""
            Me.txtYrBuilt.Text = ""
            Me.txtYrsOwned.Text = ""
            Me.txtHomeValue.Text = ""
            Me.rtbSpecialInstructions.Text = ""
            Me.pnlCustomerHistory.Controls.Clear()
            Me.Text = "Sales Department"

            Try
                Dim ctrl
                For Each ctrl In Me.pnlAFPics.Controls
                    If TypeOf (ctrl) Is ListView Then
                        ctrl.items.clear()
                    End If

                Next
            Catch ex As Exception
                Dim ctrl
                For Each ctrl In Me.pnlAFPics.Controls
                    If TypeOf (ctrl) Is ListView Then
                        ctrl.items.clear()
                    End If

                Next
            End Try

            Exit Sub
        End If

        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetCustomerINFO", Cnn)

        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        Try
            Cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            While r1.Read
                Me.txtContact1.Text = r1.Item(5) & " " & r1.Item(6)
                Me.txtContact2.Text = r1.Item(7) & " " & r1.Item(8)
                Me.txtAddress.Text = r1.Item(9) & " " & vbCrLf & r1.Item(10) & ", " & r1.Item(11) & " " & r1.Item(12)
                If r1.Item(7) = "" Then
                    Me.txtWorkHours.Text = r1.Item(5) & ": " & r1.Item(19)
                Else
                    Me.txtWorkHours.Text = r1.Item(5) & ": " & r1.Item(19) & vbNewLine & (r1.Item(7) & ": " & r1.Item(20))
                End If
                Me.txtHousePhone.Text = r1.Item(13)
                Me.txtaltphone1.Text = r1.Item(14)
                Me.txtaltphone2.Text = r1.Item(16)
                Me.txtAlt1Type.Text = r1.Item(15)
                Me.txtAlt2Type.Text = r1.Item(17)
                Me.lnkEmail.Text = r1.Item(52)
                Dim d
                d = Split(r1.Item(29), " ", 2)
                Trim(d(0))
                Me.txtApptDate.Text = d(0)
                Me.txtApptDay.Text = r1.Item(30)
                '' r1.item(31) = 'ApptTime' Field | (DateTime, null) Type on table [ EnterLead ] 
                '' sample data from 68336:> '1900-01-01 15:00:00.000'
                '' im pretty sure what is trying to happen here is just to pull out the 
                '' the time like "5:00 PM" and its getting lost in translation. 



                'Dim t = Split(r1.Item(31), " ", 2)
                'Dim u = t(1)
                'Dim u2 As String
                'Dim u3 As String
                'If u.Length = 11 Then
                '    u2 = Microsoft.VisualBasic.Left(u, 5)
                '    u3 = Microsoft.VisualBasic.Right(u, 3)
                '    u = u2 & u3
                'Else
                '    u2 = Microsoft.VisualBasic.Left(u, 4)
                '    u3 = Microsoft.VisualBasic.Right(u, 3)
                '    u = u2 & u3
                'End If

                Dim _Hour As Object = r1.Item("ApptTime").ToString
                Dim dateTime() = Split(_Hour, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                Dim _date As String = ""
                Dim _time As String = ""
                Dim _AmPM As String = ""
                _date = dateTime(0) '' 1900-01-01
                _time = dateTime(1) '' 12:00:00 
                _AmPM = dateTime(2) '' AM/PM
                Dim splitTime() = Split(_time, ":", -1, Microsoft.VisualBasic.CompareMethod.Text)
                Dim hour As String = splitTime(0) '' 12
                Dim minute As String = splitTime(1) '' 00-59 
                Dim correctedTime As String = ((hour & ":" & minute) & " " & _AmPM)
                Me.txtApptTime.Text = correctedTime.ToString
                Me.txtProducts.Text = r1.Item(21) & vbCrLf & r1.Item(22) & vbCrLf & r1.Item(23)
                Me.txtColor.Text = r1.Item(24)
                Me.txtQty.Text = r1.Item(25)
                Me.txtYrBuilt.Text = r1.Item(27)
                Me.txtYrsOwned.Text = r1.Item(26)
                Me.txtHomeValue.Text = r1.Item(28)
                Me.rtbSpecialInstructions.Text = r1.Item(32)
            End While
            r1.Close()
            Cnn.Close()

        Catch ex As Exception
            'Cnn.Close()
            'Me.PullCustomerINFO(ID)
            MsgBox("Lost Network Connection! Pull Customer Info" & ex.ToString, MsgBoxStyle.Critical, "Server not Available")
        End Try
        If Me.txtHousePhone.Text <> "" Then
            Me.btnMain.Text = "Call Main Phone - " & Me.txtHousePhone.Text
            Me.btnMain.Visible = True
        End If
        If Me.txtaltphone1.Text <> "" Then
            Me.btnAltPhone1.Text = "Call Alt Phone 1 - " & Me.txtaltphone1.Text
            Me.btnAltPhone1.Visible = True
        End If
        If Me.txtaltphone2.Text <> "" Then
            Me.btnAltPhone2.Text = "Call Alt Phone 2 - " & Me.txtaltphone2.Text
            Me.btnAltPhone2.Visible = True
        End If
        ''Populate current reps

        Dim cnn2 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet1 As SqlCommand
        Dim r2 As SqlDataReader
        cmdGet1 = New SqlCommand("Select Rep1, Rep2 from enterlead where ID = " & Me.ID, cnn2)
        cmdGet1.CommandType = CommandType.Text
        cnn2.Open()
        R2 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
        r2.Read()
        '' Loads Names of Latest Rep from sales rep pull list and will also add 
        '' the name to rep combos if they are not part of the current rep list 
        Try
            Me.cboRep1.Text = r2.Item(0)
            Me.ToolStripComboBox1.Text = r2.Item(0)
            If Me.cboRep1.Text = "" And r2.Item(0) <> "" Then
                Me.cboRep1.Items.Add(r2.Item(0))
                Me.cboRep2.Items.Add(r2.Item(0))
                Me.ToolStripComboBox1.Items.Add(r2.Item(0))
                Me.ToolStripComboBox2.Items.Add(r2.Item(0))
                Me.ToolStripComboBox1.Text = r2.Item(0)
                Me.cboRep1.Text = r2.Item(0)

            End If
        Catch ex As Exception
            Me.cboRep1.Text = ""
            Me.ToolStripComboBox1.Text = ""
        End Try
        Try
            Me.cboRep2.Text = r2.Item(1)
            Me.ToolStripComboBox2.Text = r2.Item(1)
            If Me.cboRep2.Text = "" And r2.Item(1) <> "" Then
                Me.cboRep2.Items.Add(r2.Item(1))
                Me.cboRep1.Items.Add(r2.Item(1))
                Me.ToolStripComboBox1.Items.Add(r2.Item(1))
                Me.ToolStripComboBox2.Items.Add(r2.Item(1))
                Me.ToolStripComboBox2.Text = r2.Item(1)
                Me.cboRep2.Text = r2.Item(1)

            End If
        Catch ex As Exception
            Me.cboRep2.Text = ""

        End Try

        r2.Close()
        cnn2.Close()

        Dim c As New CustomerHistory
        If ID <> "" Then
            If Me.tbMain.SelectedIndex = 1 Then
                Me.Text = "Sales Department Record ID: " & ID
            End If
            If Me.pnlCustomerHistory.Visible = True Then
                c.SetUp(Me, ID, Me.TScboCustomerHistory)
            Else
                Me.GetRidOfOldAndPutNew()
            End If

        End If

    End Sub
    Private Sub DisplayColumn(ByVal Column As String, ByVal LV As ListView)
        Dim x As Integer = 0
        Dim Y As Integer = 0
        Dim Z As Integer = 0
        If LV Is Me.lvMemorized Then
            x = 1
            Y = 45
            Z = 33
        End If

        Select Case Column
            Case Is = "Contact(s)"

                LV.Columns(0 + x).DisplayIndex = 1 + x
                LV.Columns(1 + x).DisplayIndex = 0 + x
                LV.Columns(2 + x).DisplayIndex = 2 + x
                LV.Columns(3 + x).DisplayIndex = 3 + x
                LV.Columns(4 + x).DisplayIndex = 4 + x
                LV.Columns(5 + x).DisplayIndex = 5 + x
                LV.Columns(6 + x).DisplayIndex = 6 + x
                LV.Columns(0 + x).Width = 60
                LV.Columns(1 + x).Width = 218 - Y
                LV.Columns(2 + x).Width = 255
                LV.Columns(3 + x).Width = 101
                LV.Columns(4 + x).Width = 82
                LV.Columns(5 + x).Width = 144
                LV.Columns(6 + x).Width = 99
                LV.Refresh()
            Case Is = "House Phone"
                LV.Columns(0 + x).DisplayIndex = 1 + x
                LV.Columns(1 + x).DisplayIndex = 2 + x
                LV.Columns(2 + x).DisplayIndex = 3 + x
                LV.Columns(3 + x).DisplayIndex = 0 + x
                LV.Columns(4 + x).DisplayIndex = 4 + x
                LV.Columns(5 + x).DisplayIndex = 5 + x
                LV.Columns(6 + x).DisplayIndex = 6 + x
                LV.Columns(0 + x).Width = 60
                LV.Columns(1 + x).Width = 160
                LV.Columns(2 + x).Width = 255
                LV.Columns(3 + x).Width = 218 - Y
                LV.Columns(4 + x).Width = 82
                LV.Columns(5 + x).Width = 144
                LV.Columns(6 + x).Width = 99
                LV.Refresh()

            Case Is = "Products"
                LV.Columns(0 + x).DisplayIndex = 1 + x
                LV.Columns(1 + x).DisplayIndex = 2 + x
                LV.Columns(2 + x).DisplayIndex = 3 + x
                LV.Columns(3 + x).DisplayIndex = 4 + x
                LV.Columns(4 + x).DisplayIndex = 0 + x
                LV.Columns(5 + x).DisplayIndex = 5 + x
                LV.Columns(6 + x).DisplayIndex = 6 + x
                LV.Columns(0 + x).Width = 60
                LV.Columns(1 + x).Width = 160
                LV.Columns(2 + x).Width = 255 - Y
                LV.Columns(3 + x).Width = 101
                LV.Columns(4 + x).Width = 218
                LV.Columns(5 + x).Width = 144
                LV.Columns(6 + x).Width = 99
                LV.Refresh()
            Case Is = "Appt. Date/Time"
                LV.Columns(0 + x).DisplayIndex = 1 + x
                LV.Columns(1 + x).DisplayIndex = 2 + x
                LV.Columns(2 + x).DisplayIndex = 3 + x
                LV.Columns(3 + x).DisplayIndex = 4 + x
                LV.Columns(4 + x).DisplayIndex = 5 + x
                LV.Columns(5 + x).DisplayIndex = 0 + x
                LV.Columns(6 + x).DisplayIndex = 6 + x
                LV.Columns(0 + x).Width = 60 + x
                LV.Columns(1 + x).Width = 160
                LV.Columns(2 + x).Width = 255 - Y
                LV.Columns(3 + x).Width = 101
                LV.Columns(4 + x).Width = 82
                LV.Columns(5 + x).Width = 218
                LV.Columns(6 + x).Width = 99
                LV.Refresh()

            Case Is = "Rep(s)"
                LV.Columns(0 + x + x).DisplayIndex = 1 + x
                LV.Columns(1 + x).DisplayIndex = 2 + x
                LV.Columns(2 + x).DisplayIndex = 3 + x
                LV.Columns(3 + x).DisplayIndex = 4 + x
                LV.Columns(4 + x).DisplayIndex = 5 + x
                LV.Columns(5 + x).DisplayIndex = 6 + x
                LV.Columns(6 + x).DisplayIndex = 0 + x
                LV.Columns(0 + x).Width = 60
                LV.Columns(1 + x).Width = 160
                LV.Columns(2 + x).Width = 255 - Y
                LV.Columns(3 + x).Width = 101
                LV.Columns(4 + x).Width = 82
                LV.Columns(5 + x).Width = 144
                LV.Columns(6 + x).Width = 218
                LV.Refresh()

            Case Else
                LV.Columns(0 + x).DisplayIndex = 0 + x
                LV.Columns(1 + x).DisplayIndex = 1 + x
                LV.Columns(2 + x).DisplayIndex = 2 + x
                LV.Columns(3 + x).DisplayIndex = 3 + x
                LV.Columns(4 + x).DisplayIndex = 4 + x
                LV.Columns(5 + x).DisplayIndex = 5 + x
                LV.Columns(6 + x).DisplayIndex = 6 + x
                LV.Columns(0 + x).Width = 60
                LV.Columns(1 + x).Width = 160 - Z
                LV.Columns(2 + x).Width = 255
                LV.Columns(3 + x).Width = 101
                LV.Columns(4 + x).Width = 82
                LV.Columns(5 + x).Width = 144
                LV.Columns(6 + x).Width = 99
                LV.Refresh()


        End Select

    End Sub
    Private Sub LoadReps()
        '' Load Current Rep List 
        Dim cnn1 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet2 As SqlCommand
        Dim r As SqlDataReader
        cmdGet2 = New SqlCommand("dbo.GetSalesReps", cnn1)
        cmdGet2.CommandType = CommandType.StoredProcedure

        cnn1.Open()
        r = cmdGet2.ExecuteReader(CommandBehavior.CloseConnection)
        Me.cboRep1.Items.Clear()
        Me.cboRep2.Items.Clear()
        Me.ToolStripComboBox1.Items.Clear()
        Me.ToolStripComboBox2.Items.Clear()
        Me.cboRep1.Items.Add("")
        Me.cboRep2.Items.Add("")
        Me.ToolStripComboBox1.Items.Add("")
        Me.ToolStripComboBox2.Items.Add("")
        While r.Read
            Me.cboRep1.Items.Add(r.Item(0) & " " & r.Item(1))
            Me.cboRep2.Items.Add(r.Item(0) & " " & r.Item(1))
            Me.ToolStripComboBox1.Items.Add(r.Item(0) & " " & r.Item(1))
            Me.ToolStripComboBox2.Items.Add(r.Item(0) & " " & r.Item(1))
        End While
        r.Close()
        cnn1.Close()
    End Sub
    Private Sub LoadGroups()
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdIns As SqlCommand = New SqlCommand("Select GroupName from MemorizedGroupPull Where Department = '" & Me.Name & "'", cnn)
        cmdIns.CommandType = CommandType.Text

        Me.cboFilterGroups.Items.Clear()
        Me.cboFilterGroups.Items.Add("")
        cnn.Open()
        Dim R1 As SqlDataReader
        R1 = cmdIns.ExecuteReader(CommandBehavior.CloseConnection)
        While R1.Read()
            Me.cboFilterGroups.Items.Add(R1.Item(0))
        End While
        R1.Close()
        cnn.Close()
        Me.cboFilterGroups.Items.Add("No Group Assigned")
    End Sub
    Private Sub MemorizedGroupBy()
        Dim cnn1 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet1 As SqlCommand
        Dim r1 As SqlDataReader
        Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param2 As SqlParameter = New SqlParameter("@Group", Me.cboFilterGroups.Text)
        Dim param3 As SqlParameter = New SqlParameter("@Groupby", Me.cboGroupByMemorized.Text)
        cmdGet1 = New SqlCommand("dbo.SalesDepartmentMemorizedGroupBy", cnn1)
        cmdGet1.CommandType = CommandType.StoredProcedure
        cmdGet1.Parameters.Add(param1)
        cmdGet1.Parameters.Add(param2)
        cmdGet1.Parameters.Add(param3)
        cnn1.Open()
        r1 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
        While r1.Read
            Me.lvMemorized.Groups.Add(r1.Item(0), r1.Item(0))

        End While
        Me.lvMemorized.Groups.Add("No Group Assigned", "No Group Assigned")
        r1.Close()
        cnn1.Close()
    End Sub
    Public Sub PopulateMemorized()
        If Me.lvMemorized.SelectedItems.Count <> 0 Then
            Me.lvMemorized.Tag = Me.lvMemorized.SelectedItems(0).Tag
        Else
            Me.lvMemorized.Tag = ""
        End If
        Me.lvMemorized.Groups.Clear()
        Me.lvMemorized.Items.Clear()
        If Me.cboFilterGroups.Text = "" Then
            Dim cnn1 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdGet1 As SqlCommand
            Dim r1 As SqlDataReader
            cmdGet1 = New SqlCommand("Select Distinct (Groups) from memorizedappts where Form = 'Sales' and userloggedon = '" & STATIC_VARIABLES.CurrentUser & "'", cnn1)
            cmdGet1.CommandType = CommandType.Text
            cnn1.Open()
            r1 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
            While r1.Read
                Me.lvMemorized.Groups.Add(r1.Item(0), r1.Item(0))
            End While
            Me.lvMemorized.Groups.Add("No Group Assigned", "No Group Assigned")
            r1.Close()
            cnn1.Close()

        End If
        If Me.cboGroupByMemorized.Text = "" Then

        Else
            Me.MemorizedGroupBy()
        End If
        Dim grp As String
        If Me.cboFilterGroups.Text = "" Or Me.cboFilterGroups.Text = Nothing Then
            grp = "%"
        ElseIf Me.cboFilterGroups.SelectedItem = "No Group Assigned" Then
            grp = ""
        Else
            grp = Me.cboFilterGroups.Text
        End If
        Dim x As String
        If Me.lnkOrderbyMem.Text = "Order By Appt. Date" Then
            x = "ID"
        Else
            x = "ApptDate"
        End If
        Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param2 As SqlParameter = New SqlParameter("@Group", grp)
        Dim param3 As SqlParameter = New SqlParameter("@Sortby", x)

        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand
        Dim r As SqlDataReader
        cmdGet = New SqlCommand("dbo.GetMemorizedSales", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cnn.Open()
        r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        Dim cnt As Integer = 0
        While r.Read
            Dim lv As New ListViewItem
            lv.Tag = r.Item(0).ToString
            If r.Item(23) <> "" Then
                lv.ToolTipText = r.Item(23)
                lv.ImageIndex = 8
            End If
            lv.SubItems.Add(r.Item(0))
            If r.Item(3) = "" And r.Item(4) = "" Then
                lv.SubItems.Add(r.Item(2) & ", " & r.Item(1))
            ElseIf r.Item(3) <> "" And r.Item(4) <> "" And r.Item(2) = r.Item(4) Then
                lv.SubItems.Add(r.Item(2) & ", " & r.Item(1) & " & " & r.Item(3))
            ElseIf r.Item(3) <> "" And r.Item(4) <> "" And r.Item(2) <> r.Item(4) Then
                lv.SubItems.Add(r.Item(2) & ", " & r.Item(1) & " & " & r.Item(4) & ", " & r.Item(3))
            End If
            lv.SubItems.Add(r.Item(5) & " " & r.Item(6) & ", " & r.Item(7) & " " & r.Item(8))
            lv.SubItems.Add(r.Item(17))
            If r.Item(10) = "" And r.Item(11) = "" Then
                lv.SubItems.Add(r.Item(9))
            ElseIf r.Item(10) <> "" And r.Item(11) = "" Then
                lv.SubItems.Add(r.Item(9) & " - " & r.Item(10))
            ElseIf r.Item(10) = "" And r.Item(11) <> "" Then
                lv.SubItems.Add(r.Item(9) & " - " & r.Item(11))
            ElseIf r.Item(10) <> "" And r.Item(11) <> "" Then
                lv.SubItems.Add(r.Item(9) & " - " & r.Item(10) & " - " & r.Item(11))
            End If
            Dim u As String = r.Item(16).ToString
            Dim w = InStr(u, " ")
            u = Microsoft.VisualBasic.Right(u, w + 2)
            Trim(u)
            Dim u2 As String
            Dim u3 As String
            If u.Length = 11 Then
                u2 = Microsoft.VisualBasic.Left(u, 5)
                u3 = Microsoft.VisualBasic.Right(u, 3)
                u = u2 & u3
            Else
                u2 = Microsoft.VisualBasic.Left(u, 4)
                u3 = Microsoft.VisualBasic.Right(u, 3)
                u = u2 & u3
            End If
            Dim ApptTime As String = u
            Dim ApptDate As Date = r.Item(15)
            ApptDate.ToShortDateString()
            lv.SubItems.Add(ApptDate & " - " & ApptTime)
            Dim Rep1 As String
            Dim Rep2 As String
            Try
                Rep1 = r.Item(18)
            Catch ex As Exception
                Rep1 = ""
            End Try
            Try
                Rep2 = r.Item(19)
            Catch ex As Exception
                Rep2 = ""
            End Try
            If Rep1 <> "" And Rep2 <> "" Then
                lv.SubItems.Add(Rep1 & " - " & Rep2)
            ElseIf Rep1 <> "" And Rep2 = "" Then
                lv.SubItems.Add(Rep1)
            ElseIf Rep1 = "" And Rep2 <> "" Then
                lv.SubItems.Add(Rep2)
            ElseIf Rep1 = "" And Rep2 = "" Then
                lv.SubItems.Add("")
            End If
            If Me.cboGroupByMemorized.Text = "Primary Product" Then
                If r.Item(9) <> "" Then
                    Me.lvMemorized.Groups.Add(r.Item(9), r.Item(9))
                    lv.Group = Me.lvMemorized.Groups(r.Item(9))
                Else
                    Me.lvMemorized.Groups.Add("", "")
                    lv.Group = Me.lvMemorized.Groups("")
                End If
            ElseIf Me.cboGroupByMemorized.Text = "Appt. Date" Then
                Me.lvMemorized.Groups.Add(ApptDate, ApptDate)
                lv.Group = Me.lvMemorized.Groups(ApptDate)
            ElseIf Me.cboGroupByMemorized.Text = "Sales Result" Then
                If Trim(r.Item(21)) = "" Then
                    Me.lvMemorized.Groups.Add("No Result", "No Result")
                    lv.Group = Me.lvMemorized.Groups("No Result")
                Else
                    Me.lvMemorized.Groups.Add(r.Item(21), r.Item(21))
                    lv.Group = Me.lvMemorized.Groups(r.Item(21))
                End If
            ElseIf Me.cboGroupByMemorized.Text = "Marketing Result" Then
                If Trim(r.Item(22)) = "" Then
                    Me.lvMemorized.Groups.Add("No Result", "No Result")
                    lv.Group = Me.lvMemorized.Groups("No Result")
                Else
                    Me.lvMemorized.Groups.Add(r.Item(22), r.Item(22))
                    lv.Group = Me.lvMemorized.Groups(r.Item(22))
                End If
            ElseIf Me.cboGroupByMemorized.Text = "Sales Rep" Then
                If Trim(r.Item(18)) = "" Then
                    Me.lvMemorized.Groups.Add("No Sales Rep Assigned", "No Sales Rep Assigned")
                    lv.Group = Me.lvMemorized.Groups("No Sales Rep Assigned")
                Else
                    Me.lvMemorized.Groups.Add(r.Item(18), r.Item(18))
                    lv.Group = Me.lvMemorized.Groups(r.Item(18))
                End If
            ElseIf Me.cboGroupByMemorized.Text = "City, State" Then
                Me.lvMemorized.Groups.Add(r.Item(6) & ", " & r.Item(7), r.Item(6) & ", " & r.Item(7))
                lv.Group = Me.lvMemorized.Groups(r.Item(6) & ", " & r.Item(7))
            ElseIf Me.cboGroupByMemorized.Text = "Zip Code" Then
                If Trim(r.Item(8)) = "" Then
                    Me.lvMemorized.Groups.Add("No Zip Code", "No Zip Code")
                    lv.Group = Me.lvMemorized.Groups("No Zip Code")
                Else
                    Me.lvMemorized.Groups.Add(r.Item(8), r.Item(8))
                    lv.Group = Me.lvMemorized.Groups(r.Item(8))
                End If
            Else
                If r.Item(24) = "" Then
                    lv.Group = Me.lvMemorized.Groups("No Group Assigned")
                Else
                    lv.Group = Me.lvMemorized.Groups(r.Item(24))
                End If
            End If
            lv.SubItems.Add(r.Item(20))
            Me.lvMemorized.Items.Add(lv)
            If lv.Tag = Me.lvMemorized.Tag Then
                lv.Selected = True
                lv.EnsureVisible()
            End If
        End While
        r.Close()
        cnn.Close()
        If Me.lvMemorized.SelectedItems.Count = 0 And Me.lvMemorized.Items.Count <> 0 Then
            Me.lvMemorized.TopItem.Selected = True
        End If
        If Me.lvMemorized.Items.Count = 0 Then
            Me.PullInfo("")
        End If


    End Sub
    Private Sub RefreshSelectedItem(ByVal Listview As ListView, ByVal ID As String)
        Dim cnn1 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim x As Integer = 0
        If Listview Is Me.lvMemorized Then
            x = 1
        End If
        If ID = "" Then
            Exit Sub
        End If
        Dim cmdGet1 As SqlCommand
        Dim r As SqlDataReader
        cmdGet1 = New SqlCommand("Select Id, Contact1FirstName, Contact1Lastname, Contact2FirstName, Contact2LastName,StAddress, City, State, Zip, Product1, Product2, Product3,Productacro1, Productacro2,Productacro3, ApptDate, ApptTime, HousePhone, Rep1, Rep2, NeedsSaleResult, Result, MarketingResults from enterlead where id = '" & ID & "'", cnn1)
        cmdGet1.CommandType = CommandType.Text
        cnn1.Open()
        r = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
        r.Read()
        Dim lv As ListViewItem = Listview.SelectedItems(0)
        If r.Item(3) = "" And r.Item(4) = "" Then
            lv.SubItems(1 + x).Text = (r.Item(2) & ", " & r.Item(1))
        ElseIf r.Item(3) <> "" And r.Item(4) <> "" And r.Item(2) = r.Item(4) Then
            lv.SubItems(1 + x).Text = (r.Item(2) & ", " & r.Item(1) & " & " & r.Item(3))
        ElseIf r.Item(3) <> "" And r.Item(4) <> "" And r.Item(2) <> r.Item(4) Then
            lv.SubItems(1 + x).Text = (r.Item(2) & ", " & r.Item(1) & " & " & r.Item(4) & ", " & r.Item(3))
        End If


        lv.SubItems(2 + x).Text = (r.Item(5) & " " & r.Item(6) & ", " & r.Item(7) & " " & r.Item(8))
        lv.SubItems(3 + x).Text = (r.Item(17))
        If r.Item(10) = "" And r.Item(11) = "" Then
            If r.Item(9).ToString.Length > 8 Then
                lv.SubItems(4 + x).Text = r.Item(12)
            Else
                lv.SubItems(4 + x).Text = (r.Item(9))
            End If
        ElseIf r.Item(10) <> "" And r.Item(11) = "" Then
            If (r.Item(9) & " - " & r.Item(10)).ToString.Length > 8 Then
                lv.SubItems(4 + x).Text = r.Item(12) & " - " & r.Item(13)
            Else
                lv.SubItems(4 + x).Text = r.Item(9) & " - " & r.Item(10)
            End If
        ElseIf r.Item(10) = "" And r.Item(11) <> "" Then
            If (r.Item(9) & " - " & r.Item(11)).ToString.Length > 8 Then
                lv.SubItems(4 + x).Text = r.Item(12) & " - " & r.Item(14)
            Else
                lv.SubItems(4 + x).Text = r.Item(9) & " - " & r.Item(11)
            End If
        ElseIf r.Item(10) <> "" And r.Item(11) <> "" Then
            If (r.Item(9) & " - " & r.Item(10) & " - " & r.Item(11)).ToString.Length > 8 Then
                lv.SubItems(4 + x).Text = r.Item(12) & " - " & r.Item(13) & " - " & r.Item(14)
            Else
                lv.SubItems(4 + x).Text = r.Item(9) & " - " & r.Item(10) & " - " & r.Item(11)
            End If
        End If
        Dim u As String = r.Item(16).ToString
        Dim w = InStr(u, " ")
        u = Microsoft.VisualBasic.Right(u, w + 2)
        Trim(u)
        Dim u2 As String
        Dim u3 As String
        If u.Length = 11 Then
            u2 = Microsoft.VisualBasic.Left(u, 5)
            u3 = Microsoft.VisualBasic.Right(u, 3)
            u = u2 & u3
        Else
            u2 = Microsoft.VisualBasic.Left(u, 4)
            u3 = Microsoft.VisualBasic.Right(u, 3)
            u = u2 & u3
        End If
        Dim ApptTime As String = u
        Dim ApptDate As Date = r.Item(15)
        ApptDate.ToShortDateString()
        lv.SubItems(5 + x).Text = (ApptDate & " - " & ApptTime)
        Dim Rep1 As String
        Dim Rep2 As String
        Try
            Rep1 = r.Item(18)
        Catch ex As Exception
            Rep1 = ""
        End Try
        Try
            Rep2 = r.Item(19)
        Catch ex As Exception
            Rep2 = ""
        End Try
        If Rep1 <> "" And Rep2 <> "" Then
            lv.SubItems(6 + x).Text = (Rep1 & " - " & Rep2)
        ElseIf Rep1 <> "" And Rep2 = "" Then
            lv.SubItems(6 + x).Text = (Rep1)
        ElseIf Rep1 = "" And Rep2 <> "" Then
            lv.SubItems(6 + x).Text = (Rep2)
        ElseIf Rep1 = "" And Rep2 = "" Then
            lv.SubItems(7 + x).Text = ("")
        End If
        lv.SubItems(7 + x).Text = r.Item(20)

        r.Close()
        cnn1.Close()
    End Sub
#End Region

#Region "Whole Form Events"






    Private Sub SalesDepartment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Me.Parent = Main
        Me.Text = "Sales Department"
        Me.Label4.Location = New System.Drawing.Point((Me.tpSummary.Width / 2) - 87, Me.Label4.Location.Y)
        ' Me.lvAttachedFiles.Size = New System.Drawing.Size(Me.pnlAFPics.Width / 2 - 5, Me.lvAttachedFiles.Height)
        ' Me.lvJobPics.Size = New System.Drawing.Size(Me.pnlAFPics.Width / 2 - 5, Me.lvJobPics.Height)
        Me.ButtonConfig()
        Me.LoadReps()
        '' Set Up Summary Tab 

        '' Set up Sales List Tab 
        Me.TScboCustomerHistory.Text = "All"
        Me.tscboAFPicsFilter.Text = "All"
        Me.tsCustomerLog.Visible = True
        Me.tsAFPics.Visible = False
        Me.pnlCustomerHistory.Visible = True
        Me.pnlAFPics.Visible = False



        Me.SplitContainer1.SplitterDistance = 218
        Me.SplitContainer1.IsSplitterFixed = True
        Me.btnExpandSalesList.Text = Chr(187)
        Me.cboDateRangeCustomerList.Text = "All"



        Me.ToolbarConfig(1)
        Me.PopulateNoResults()
        Me.LoadGroups()
        Dim x As New ScheduledActions
        x.SetUp(Me)




        dtp1CustomerList.Value = CDate("1/1/1900 12:00:00 AM")
        Me.cboSalesList.Text = "Unfiltered Sales Dept. List"
        Me.PopulateMemorized()
        If Me.lvnoresults.Items.Count <> 0 Then
            Me.lvnoresults.TopItem.Selected = True
        End If
        If Me.lvMemorized.Items.Count <> 0 Then
            Me.lvMemorized.TopItem.Selected = True
        End If

        ' Me.lvAttachedFiles.Size = New System.Drawing.Size(Me.pnlAFPics.Width / 2 - 5, Me.lvAttachedFiles.Height)
        'Me.lvJobPics.Size = New System.Drawing.Size(Me.pnlAFPics.Width / 2 - 5, Me.lvJobPics.Height)

        Me.cboDateRangeSummary.Text = "Last Week"
        Dim r As New Sales_Performance_Report()


        Me.LoadComplete = True
        Me.cboDateRangeSummary_SelectedIndexChanged(Nothing, Nothing)
        If Me.WindowState <> FormWindowState.Normal Then

        End If
        Me.tbMain.TabPages(0).Font = New Font("Verdana", 12, FontStyle.Strikeout)


        'MsgBox(Me.tsSalesDepartment.Width.ToString & " - " & Me.txt2CustomerList.Location.X.ToString)
    End Sub



    Private Sub btnExpandSalesList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandSalesList.Click

        Select Case Me.btnExpandSalesList.Text
            Case Chr(187)
                Me.SplitContainer1.SplitterDistance = 2500
                Me.SplitContainer1.SplitterWidth = 100
                Me.btnExpandMemorize.Text = Chr(171)
                Me.btnExpandSalesList.Text = Chr(171)
            Case Chr(171)
                Me.SplitContainer1.SplitterDistance = 218
                Me.SplitContainer1.SplitterWidth = 1
                Me.btnExpandSalesList.Text = Chr(187)
                Me.btnExpandMemorize.Text = Chr(187)
        End Select
    End Sub



  
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbMain.SelectedIndexChanged
        Dim x As Integer = 0
        x = Me.tbMain.SelectedIndex
        Select Case x
            Case 0
                Me.tbMain.TabPages(x).ImageKey = "Summary- Selected.png"
                Me.tbMain.TabPages(1).ImageKey = "Customer List.png"
                Me.tbMain.TabPages(2).ImageKey = "Issue Leads.png"
                Me.tbMain.TabPages(3).ImageKey = "References.png"
                Me.tbMain.TabPages(4).ImageKey = "Reports.png"
                Me.ToolbarConfig(1)
                Me.PopulateNoResults()
                If Me.lvnoresults.SelectedItems.Count = 0 Then
                    Me.ID = ""
                    STATIC_VARIABLES.CurrentID = Me.ID
                Else
                    Me.ID = Me.lvnoresults.SelectedItems(0).SubItems(1).Text
                    STATIC_VARIABLES.CurrentID = Me.ID
                End If

                Exit Select
            Case 1
                Me.tbMain.TabPages(x).ImageKey = "Customer List- Selected.png"
                Me.tbMain.TabPages(0).ImageKey = "Summary.png"
                Me.tbMain.TabPages(2).ImageKey = "Issue Leads.png"
                Me.tbMain.TabPages(3).ImageKey = "References.png"
                Me.tbMain.TabPages(4).ImageKey = "Reports.png"
                ' Me.lvAttachedFiles.Size = New System.Drawing.Size(Me.pnlAFPics.Width / 2 - 5, Me.lvAttachedFiles.Height)
                ' Me.lvJobPics.Size = New System.Drawing.Size(Me.pnlAFPics.Width / 2 - 5, Me.lvJobPics.Height)
                If Me.TabControl2.SelectedIndex = 0 Then
                    If Me.lvSales.Items.Count <> 0 And Me.lvSales.SelectedItems.Count = 0 Then
                        Me.lvSales.TopItem.Selected = True
                    End If
                    Me.ToolbarConfig(2)
                    If Me.lvSales.SelectedItems.Count = 0 Then
                        Me.ID = ""
                        STATIC_VARIABLES.CurrentID = Me.ID
                    Else
                        Me.ID = Me.lvSales.SelectedItems(0).Text
                        STATIC_VARIABLES.CurrentID = Me.ID
                    End If
                Else
                    Me.ToolbarConfig(3)
                    If Me.lvMemorized.SelectedItems.Count = 0 Then
                        Me.ID = ""
                        STATIC_VARIABLES.CurrentID = Me.ID
                    Else
                        Me.ID = Me.lvMemorized.SelectedItems(0).Tag
                        STATIC_VARIABLES.CurrentID = Me.ID
                    End If
                End If


                Exit Select
            Case 2
                Me.tbMain.TabPages(x).ImageKey = "Issue Leads- Selected.png"
                Me.tbMain.TabPages(0).ImageKey = "Summary.png"
                Me.tbMain.TabPages(1).ImageKey = "Customer List.png"
                Me.tbMain.TabPages(3).ImageKey = "References.png"
                Me.tbMain.TabPages(4).ImageKey = "Reports.png"
                Me.ID = ""
                STATIC_VARIABLES.CurrentID = Me.ID
                ToolbarConfig(4)

                Dim d As New Issue_Leads(True, "")

                Exit Select
            Case 3
                Me.tbMain.TabPages(x).ImageKey = "References- Selected.png"
                Me.tbMain.TabPages(0).ImageKey = "Summary.png"
                Me.tbMain.TabPages(2).ImageKey = "Issue Leads.png"
                Me.tbMain.TabPages(1).ImageKey = "Customer List.png"
                Me.tbMain.TabPages(4).ImageKey = "Reports.png"
                Me.ID = ""
                STATIC_VARIABLES.CurrentID = Me.ID
                Exit Select
            Case 4
                Me.tbMain.TabPages(x).ImageKey = "Reports- Selected.png"
                Me.tbMain.TabPages(0).ImageKey = "Summary.png"
                Me.tbMain.TabPages(2).ImageKey = "Issue Leads.png"
                Me.tbMain.TabPages(3).ImageKey = "References.png"
                Me.tbMain.TabPages(1).ImageKey = "Customer List.png"
                Me.ID = ""
                STATIC_VARIABLES.CurrentID = Me.ID
                Exit Select
        End Select
        If Me.tbMain.SelectedIndex <> 0 Then
            Dim c As Integer = Me.pnlScheduledTasks.Controls.Count
            Dim i As Integer
            For i = 1 To c
                Dim all As Panel = Me.pnlScheduledTasks.Controls(i - 1)
                If all.BorderStyle = BorderStyle.FixedSingle Then
                    all.BorderStyle = BorderStyle.None
                End If
            Next
        End If
    End Sub
#End Region

#Region "Summary Page Events"
    ''Scheduled Tasks Controls
    Private Sub btnMarkTaskAsDone_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMarkTaskAsDone.Click
        Dim counter As Integer
        Dim cnt = Me.pnlScheduledTasks.Controls.Count
        Dim i As Integer
        For i = 1 To cnt
            Dim t As Panel = Me.pnlScheduledTasks.Controls(i - 1)
            If t.BorderStyle = BorderStyle.FixedSingle Then
                counter = counter + 1
            End If
            If t.BorderStyle = BorderStyle.FixedSingle Then
                t.BorderStyle = BorderStyle.None
                Dim name = t.Name.Substring(3)
                Dim c As New ScheduledActions
                If Me.btnMarkTaskAsDone.Text = "Mark Task as Done" Then
                    c.Completed(CInt(name))
                Else
                    c.UndoCompleted(CInt(name))
                End If
                Exit For
            End If
        Next
        If counter = 0 Then
            MsgBox("You must select a Task.", MsgBoxStyle.Exclamation, "Please Select a Task")
            Exit Sub
        End If

        Me.pnlScheduledTasks.Controls.Clear()
        Dim x As New ScheduledActions
        x.SetUp(Me)
    End Sub
    Private Sub btnSAPreferences_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSAPreferences.Click
        Dim y As Integer
        Try
            y = InputBox$("Enter the number of days from Today and back you would like ""Completed Tasks"" to display.", "Enter Number of Days")

        Catch ex As Exception
            'MsgBox("This field only accepts numbers!" & vbCr & "Operation Aborted!", MsgBoxStyle.Exclamation, "You Must Enter a Number")
            Exit Sub
        End Try
        If y <= 0 Then
            'MsgBox("You must enter a number!" & vbCr & "Operation Aborted!", MsgBoxStyle.Exclamation, "Cannot Write Blank Value")
            Exit Sub
        End If
        Dim x As New ScheduledActions
        x.Preferences(y, Me)
        Me.pnlScheduledTasks.Controls.Clear()
        x.SetUp(Me)
        If Me.cboSalesList.Text = "Scheduled Tasks List" Then
            Me.cboSalesList_SelectedIndexChanged(Nothing, Nothing)
        End If
    End Sub
    Private Sub btnShowAllCompletedTasks_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShowAllCompletedTasks.Click
        Me.btnShowAllCompletedTasks.CheckState = CheckState.Checked
        Me.btnRemoveAllScheduledTask.CheckState = CheckState.Unchecked
        Dim x As New ScheduledActions
        x.ShowAll(Me)
        Me.pnlScheduledTasks.Controls.Clear()
        x.SetUp(Me)
    End Sub
    Private Sub btnRemoveAllScheduledTask_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveAllScheduledTask.Click
        Me.btnShowAllCompletedTasks.CheckState = CheckState.Unchecked
        Me.btnRemoveAllScheduledTask.CheckState = CheckState.Checked
        Dim x As New ScheduledActions
        x.HideAll(Me)
        Me.pnlScheduledTasks.Controls.Clear()
        x.SetUp(Me)
    End Sub
    Private Sub btnRemoveThisCompletedTask_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveThisCompletedTask.Click
        Me.btnRemoveThisCompletedTask.Visible = False
        Dim counter As Integer
        Dim cnt = Me.pnlScheduledTasks.Controls.Count
        Dim i As Integer
        For i = 1 To cnt
            Dim t As Panel = Me.pnlScheduledTasks.Controls(i - 1)
            If t.BorderStyle = BorderStyle.FixedSingle Then
                counter = counter + 1
            End If
            If t.BorderStyle = BorderStyle.FixedSingle Then
                t.BorderStyle = BorderStyle.None
                Dim ID = t.Name.Substring(3)
                Dim c As New ScheduledActions
                c.HideTask(CInt(ID))
                Exit For
            End If

        Next
        If counter = 0 Then
            MsgBox("You must select a Task.", MsgBoxStyle.Exclamation, "Please Select a Task")
            Exit Sub
        End If

        Me.pnlScheduledTasks.Controls.Clear()
        Dim x As New ScheduledActions
        x.SetUp(Me)
    End Sub
    ''Sales Result Controls
    Private Sub btnSalesResult_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSalesResult.Click
        If Me.ID = "" Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        SDResult.ID = Me.ID
        SDResult.ShowDialog()
    End Sub

    Private Sub lvnoresults_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvnoresults.DoubleClick
        If Me.lvnoresults.SelectedItems.Count = 0 Then
            Exit Sub
        End If

        Me.tbMain.SelectedIndex = 1
        Me.TabControl2.SelectedIndex = 0
        Me.cboSalesList.Text = "Records Needing Sales Results"
        Me.txtSingleRecordInput.Text = Me.lvnoresults.SelectedItems(0).SubItems(1).Text
        Me.btnSingleRecord_Click(Nothing, Nothing)


    End Sub
    Private Sub lvnoresults_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvnoresults.SelectedIndexChanged
        If Me.lvnoresults.SelectedItems.Count = 0 Then
            Me.ID = ""
            STATIC_VARIABLES.CurrentID = Me.ID
            Exit Sub
        End If

        Me.ID = Me.lvnoresults.SelectedItems(0).SubItems(1).Text
        STATIC_VARIABLES.CurrentID = Me.ID
    End Sub

#End Region
#Region "Customer List Page Events"
#Region "Customer List Page Toolbar Buttons"
    Private Sub btnSalesResult2_ButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSalesResult2.ButtonClick
        If Me.ID = "" Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        SDResult.ID = Me.ID
        SDResult.ShowDialog()
    End Sub
    Private Sub btnBuildList_ButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuildList.ButtonClick
        SalesListBuilder.Rollback = Me.cboSalesList.Text
        If Me.cboSalesList.Text <> "Custom..." Then
            Me.cboSalesList.Text = "Custom..."
        Else
            SalesListBuilder.ShowDialog()
        End If
    End Sub



#Region "Date Range Controls"
    Private Sub txt2CustomerList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt2CustomerList.Click
        Me.dtp2CustomerList.Value = Today
        Me.dtp2CustomerList.Focus()
        txt2CustomerList.Visible = False

        txt2CustomerList.Text = Me.dtp2CustomerList.Value.ToShortDateString
        Me.cboDateRangeCustomerList.Text = "Custom"
    End Sub
    Private Sub txt1CustomerList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1CustomerList.Click
        txt1CustomerList.Visible = False
        Me.dtp1CustomerList.Value = Today
        dtp1CustomerList.Focus()
        txt1CustomerList.Text = dtp1CustomerList.Value.ToShortDateString
        Me.cboDateRangeCustomerList.Text = "Custom"
    End Sub
    Private Sub dtp2CustomerList_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp2CustomerList.GotFocus
        'txt1CustomerList.Visible = False
        txt2CustomerList.Text = Me.dtp2CustomerList.Value.ToShortDateString
        txt2CustomerList.Visible = False
        Me.cboDateRangeCustomerList.Text = "Custom"
    End Sub
    Private Sub dtp2CustomerList_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp2CustomerList.LostFocus
        If Me.LastD1 = txt1CustomerList.Text And Me.LastD2 = txt2CustomerList.Text Then
            Exit Sub
        End If
        Dim c As New SalesListManager
    End Sub
    Private Sub dtp2CustomerList_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp2CustomerList.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        txt2CustomerList.Text = Me.dtp2CustomerList.Value.ToShortDateString
    End Sub
    Private Sub dtp1CustomerList_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp1CustomerList.GotFocus
        txt1CustomerList.Visible = False
        txt1CustomerList.Text = dtp1CustomerList.Value.ToShortDateString
        'txt2CustomerList.Visible = False
        Me.cboDateRangeCustomerList.Text = "Custom"
    End Sub
    Private Sub dtp1CustomerList_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp1CustomerList.LostFocus
        If txt1CustomerList.Text = Me.LastD1 And txt2CustomerList.Text = Me.LastD2 Then
            Exit Sub
        End If
        Dim c As New SalesListManager
    End Sub
    Private Sub dtp1CustomerList_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp1CustomerList.ValueChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        txt1CustomerList.Text = dtp1CustomerList.Value.ToShortDateString
    End Sub
#End Region
    Private Sub cboDateRangeCustomerList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDateRangeCustomerList.SelectedIndexChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
        If Me.cboDateRangeCustomerList.Text = "All" Then
            txt1CustomerList.Text = ""
            txt1CustomerList.Visible = True
            txt2CustomerList.Text = ""
            txt2CustomerList.Visible = True

        End If

        If Me.cboDateRangeCustomerList.Text <> "All" And Me.cboDateRangeCustomerList.Text <> "Custom" Then
            Dim d As New DTPManipulation(Me.cboDateRangeCustomerList.Text)
            dtp1CustomerList.Value = d.retDateFrom
            Me.dtp2CustomerList.Value = d.retDateTo
            txt1CustomerList.Visible = False
            txt2CustomerList.Visible = False
            txt1CustomerList.Text = d.retDateFrom.ToString
            txt2CustomerList.Text = d.retDateTo.ToString
        End If
        If Me.LastD1 = txt1CustomerList.Text And Me.LastD2 = txt2CustomerList.Text Then
            Exit Sub
        End If
        Dim c As New SalesListManager
    End Sub

#End Region

#Region "Customer List Controls"

  

    

    Private Sub lblGroupBy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblGroupBy.Click
        Me.cboGroupSales.Focus()
        Me.cboGroupSales.DroppedDown = True
    End Sub
    Private Sub cboGroupSales_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGroupSales.SelectedIndexChanged
        If Me.cboGroupSales.Text <> "" Then
            Me.lblGroupBy.Text = Me.cboGroupSales.Text
            Me.lblGroupBy.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblGroupBy.ForeColor = Color.Black
        Else
            Me.lblGroupBy.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblGroupBy.ForeColor = Color.Gray
            Me.lblGroupBy.Text = "Group By"
        End If
        Dim c As New SalesListManager
    End Sub




    Private Sub cboDisplayColumn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDisplayColumn.SelectedIndexChanged
        If Me.cboDisplayColumn.Text <> "" Then
            Me.lblDisplayColumn.Text = Me.cboDisplayColumn.Text
            Me.lblDisplayColumn.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblDisplayColumn.ForeColor = Color.Black
        Else
            Me.lblDisplayColumn.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblDisplayColumn.ForeColor = Color.Gray
            Me.lblDisplayColumn.Text = "Change Display Column"
        End If
        Me.DisplayColumn(Me.cboDisplayColumn.Text, Me.lvSales)
    End Sub




    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        If Me.LinkLabel2.Text = "Order By ID" Then
            Me.lvSales.Sorting = Windows.Forms.SortOrder.Descending
            Me.LinkLabel2.Text = "Order By Appt. Date"
            Me.LinkLabel2.Location = New System.Drawing.Point(93, 3)
            Me.lvSales.Refresh()
        Else
            Me.lvSales.Sorting = Windows.Forms.SortOrder.None
            Me.LinkLabel2.Text = "Order By ID"
            Me.LinkLabel2.Location = New System.Drawing.Point(114, 3)
            Dim c As New SalesListManager
        End If
    End Sub
    Private Sub cboSalesList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSalesList.SelectedIndexChanged
        If Me.cboSalesList.Text = "Unconfirmed Appts. For Today" Then
            cboDateRangeCustomerList.Text = "Today"
        ElseIf Me.cboSalesList.Text = "Custom..." Then
            SalesListBuilder.ShowDialog()
            Exit Sub
        ElseIf Me.cboSalesList.Text = "Issue Leads List" Then
            If Me.dtpIssueLeads.Value = Today Then
                Me.cboDateRangeCustomerList.SelectedItem = "Today"
            Else
                Me.dtp1CustomerList.Value = Me.dtpIssueLeads.Value
                Me.dtp2CustomerList.Value = Me.dtpIssueLeads.Value
                Me.cboDateRangeCustomerList.SelectedItem = "Custom"
            End If
        End If

        Dim c As New SalesListManager
    End Sub
    Public Sub lvSales_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvSales.SelectedIndexChanged
        If Me.lvSales.SelectedItems.Count <> 0 Then
            Me.RefreshSelectedItem(Me.lvSales, Me.lvSales.SelectedItems(0).Text)
            Me.PullInfo(Me.lvSales.SelectedItems(0).Text)
            Me.ID = Me.lvSales.SelectedItems(0).Text
            STATIC_VARIABLES.CurrentID = Me.ID
        Else
            Me.PullInfo("")
            Me.ID = ""
            STATIC_VARIABLES.CurrentID = Me.ID
        End If

    End Sub
#End Region

#Region "Customer Log & Attached Files"
    Private Sub TScboCustomerHistory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TScboCustomerHistory.SelectedIndexChanged
        If Me.lvSales.SelectedItems.Count <> 0 Then
            Dim c As New CustomerHistory
            c.SetUp(Me, Me.lvSales.SelectedItems(0).Text, Me.TScboCustomerHistory)
        End If
    End Sub

    Private Sub tsbtnAFPics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtnAFPics.Click

        Me.tsCustomerLog.Visible = False
        Me.tsAFPics.Visible = True
        Me.pnlCustomerHistory.Visible = False
        Me.pnlAFPics.Visible = True
        Me.tsAttachedFilesNAV.Enabled = True

        ''Dim x As New PopulateAF(Me)
        'Me.GetRidOfOldAndPutNew()
        'Dim InitPoint As System.Drawing.Point = New System.Drawing.Point(0, 0)
        'Dim InitPoint2 As System.Drawing.Point = New System.Drawing.Point(364, 0)

        '' edit 8-29-15 
        '' need a calculation here to determine size of panel the control(s) are going into
        '' and to size the controls accordingly with a 10-15 px buffer per andy
        '' 
        '' 
        Dim widthOfParent As Integer = pnlAFPics.Width
        Dim widthOfControl As Integer = (widthOfParent / 2) - 20
        Dim heightOfParent As Integer = pnlAFPics.Height
        Dim heightOfControl As Integer = (heightOfParent - 30)
        Dim InitPoint As System.Drawing.Point = New System.Drawing.Point((0 + 10), (0 + 10))
        Dim InitPoint2 As System.Drawing.Point = New System.Drawing.Point(((widthOfParent / 2) + 10), (0 + 10))

        Dim rc As ReusableListViewControl = New ReusableListViewControl
        rc.GenerateListControl(pnlAFPics, (STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID).ToString, InitPoint, "lsAF", heightOfControl, widthOfControl)
        Dim rc2 As ReusableListViewControl = New ReusableListViewControl
        rc2.GenerateListControl(pnlAFPics, (STATIC_VARIABLES.JobPicturesFileDirectory & STATIC_VARIABLES.CurrentID).ToString, InitPoint2, "lsJP", heightOfControl, widthOfControl)

    End Sub

    Private Sub tsbtnShowCH_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbtnShowCH.Click
        Me.tsCustomerLog.Visible = True
        Me.tsAFPics.Visible = False
        Me.pnlCustomerHistory.Visible = True
        Me.pnlAFPics.Visible = False
        Me.tsAttachedFilesNAV.Enabled = False
        If Me.ID <> "" Then
            Dim x As New CustomerHistory
            x.SetUp(Me, ID, Me.TScboCustomerHistory)
        End If
    End Sub

    Private Sub tscboAFPicsFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tscboAFPicsFilter.SelectedIndexChanged
        
        Dim p_width As Integer = Me.pnlAFPics.Width
        Dim p_height As Integer = Me.pnlAFPics.Height
        ''
        '' change tslblAFPIC text to respective 'filter'
        '' 
        '' tslblAFPic
        ''

        Select Case Me.tscboAFPicsFilter.Text
            Case Is = "All"
                Me.tslblAFPic.Text = "Attached Files and Job Pictures..."
                Dim x As Control

                
                For Each x In Me.pnlAFPics.Controls
                    If x.Name.ToString = "lsAF" Then
                        Dim prevLocX As Integer = x.Location.X
                        Dim prevLocY As Integer = x.Location.Y
                        x.Size = New System.Drawing.Point((p_width / 2) - 20, p_height - 10)
                        x.Location = New Point(prevLocX + 5, 5)

                    End If
                    If x.Name.ToString = "lsJP" Then
                         
                        x.Size = New System.Drawing.Point((p_width / 2) - 20, (p_height - 10))
                        x.Location = New Point((p_width / 2) + 5, 5)
                    End If
                Next
                '        Me.lvAttachedFiles.Dock = DockStyle.Left
                '        Me.lvJobPics.Dock = DockStyle.Right
                '        Me.lvAttachedFiles.Visible = True
                '        Me.lvJobPics.Visible = True
                '        Me.tslblAFPic.Text = "Attached Files and Job Pictures..."
            Case Is = "Job Pictures"
                Me.tslblAFPic.Text = "Job Pictures"
                ''737,172 

                Dim x As Control
                For Each x In Me.pnlAFPics.Controls
                    If x.Name.ToString = "lsAF" Then
                        x.Size = New System.Drawing.Point(0, 0)
                    End If
                    If x.Name.ToString = "lsJP" Then
                        x.Size = New System.Drawing.Point((p_width - 20), (p_height - 10))
                        x.Location = New Point(5, 5)
                    End If
                Next
                '        Me.lvAttachedFiles.Visible = False
                '        Me.lvJobPics.Visible = True
                '        Me.lvJobPics.Dock = DockStyle.Fill
                '        Me.tslblAFPic.Text = "Job Pictures..."
            Case Is = "Attached Files"
                Me.tslblAFPic.Text = "Attached Files"
                Dim x As Control
                For Each x In Me.pnlAFPics.Controls
                    If x.Name.ToString = "lsAF" Then
                        x.Size = New System.Drawing.Point((p_width - 20), (p_height - 10))
                        x.Location = New Point(5, 5)

                    End If
                    If x.Name.ToString = "lsJP" Then
                        x.Size = New System.Drawing.Point(0, 0)
                    End If
                Next
                '        Me.lvJobPics.Visible = False
                '        Me.lvAttachedFiles.Visible = True
                '        Me.lvAttachedFiles.Dock = DockStyle.Fill
                '        Me.tslblAFPic.Text = "Attached Files..."
        End Select
    End Sub

    Private Sub btnLogCall_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogCall.ButtonClick

        If Me.ID = "" Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        LogPhoneCall.ID = Me.ID
        LogPhoneCall.frm = Me
        LogPhoneCall.Contact1 = Me.txtContact1.Text
        LogPhoneCall.Contact2 = Me.txtContact2.Text
        LogPhoneCall.ShowDialog()

    End Sub

    Private Sub btnCalledandCancelled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalledandCancelled.Click

        If Me.ID = "" Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        Dim s = Split(Me.txtContact1.Text, " ")
        Dim c1 = s(0)
        Dim s2 = Split(Me.txtContact2.Text, " ")
        Dim c2 = s2(0)
        CandCNotes.ID = Me.ID
        CandCNotes.Contact1 = c1
        CandCNotes.Contact2 = c2
        CandCNotes.OrigApptDate = Me.txtApptDate.Text
        CandCNotes.OrigApptTime = Me.txtApptTime.Text
        CandCNotes.frm = Me
        CandCNotes.ShowDialog()
    End Sub
#End Region
#Region "Memorized Controls"

    Private Sub lvMemorized_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMemorized.Click
        If Me.IfExists = True Then
            Me.TT.Dispose()
        End If
        Me.SelItem.ToolTipText = Nothing
        Me.SelItem.ToolTipText = Me.lvMemorized.GetItemAt(Me.xcordinate, Me.ycordinate).ToolTipText ''Come back
    End Sub

    Private Sub lvMemorized_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMemorized.DoubleClick
        Me.TT = New ToolTip
        Me.IfExists = True
        If Me.SelItem Is Nothing Then
            Exit Sub
        End If
        Dim p As System.Drawing.Point = New System.Drawing.Point(10, Me.ycordinate + 15)
        Dim notes As String
        If Me.SelItem.ToolTipText <> "" Then
            Dim c As New TruncateNotes
            c.Truncate(Me.SelItem.ToolTipText, Me.lvMemorized)
            notes = c.NewSTRING
            TT.ToolTipIcon = ToolTipIcon.Info
            TT.ToolTipTitle = "Memorized Record Notes"
            Me.TT.Show(notes, Me.lvMemorized, p, 30000)
        End If



    End Sub

    Private Sub lvMemorized_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvMemorized.MouseMove
        Me.xcordinate = e.X
        Me.ycordinate = e.Y
    End Sub
    Public Sub lvMemorized_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMemorized.SelectedIndexChanged
        If Me.lvMemorized.SelectedItems.Count <> 0 Then
            Me.RefreshSelectedItem(Me.lvMemorized, Me.lvMemorized.SelectedItems(0).Tag)
            Me.PullInfo(Me.lvMemorized.SelectedItems(0).Tag)
            Me.ID = Me.lvMemorized.SelectedItems(0).Tag
            STATIC_VARIABLES.CurrentID = Me.ID
            ''add code to refresh line item without refreshing entire list 
        Else
            Me.PullInfo("")
            Me.ID = ""
            STATIC_VARIABLES.CurrentID = Me.ID
        End If
    End Sub





#End Region
#End Region

#Region "Issue Leads"



#End Region




    Private Sub txtSingleRecordInput_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSingleRecordInput.GotFocus
        Me.txtSingleRecordInput.Text = ""
    End Sub


    Private Sub txtSingleRecordInput_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSingleRecordInput.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "") Then
            Me.btnSingleRecord.Enabled = True
        ElseIf x.ToString = vbCr Then
            Me.lvSales.Focus()
            Me.btnSingleRecord_Click(Nothing, Nothing)
            Me.btnBuildList.HideDropDown()


        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txtSingleRecordInput.Text.Length = 7 Then
            e.KeyChar = Nothing
        End If
    End Sub

    Public Sub btnSingleRecord_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSingleRecord.Click
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand
        Dim r As SqlDataReader
        cmdGet = New SqlCommand("Select Count (ID)from enterlead where ID  = " & Me.txtSingleRecordInput.Text, cnn)
        cmdGet.CommandType = CommandType.Text
        cnn.Open()
        r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r.Read()
        If r.Item(0) = 0 Then
            MsgBox("Record Doesn't Exist!", MsgBoxStyle.Exclamation, "Record # Invalid")
            Me.txtSingleRecordInput.Text = "[Enter Record ID]"
            Me.btnSingleRecord.Enabled = False
            Me.btnBuildList.HideDropDown()
            Exit Sub
        End If
        r.Close()
        cnn.Close()

        Try
            Me.lvSales.Items(Me.txtSingleRecordInput.Text).Selected = True
            Me.lvSales.SelectedItems(0).EnsureVisible()

        Catch ex As Exception
            If Me.cboSalesList.Text <> "Unfiltered Sales Dept. List" Or Me.cboDateRangeCustomerList.Text <> "All" Then
                Me.cboSalesList.Text = "Unfiltered Sales Dept. List"
                Me.cboDateRangeCustomerList.Text = "All"
            End If
            Try
                Me.lvSales.Items(Me.txtSingleRecordInput.Text).Selected = True
                Me.lvSales.SelectedItems(0).EnsureVisible()
            Catch ex2 As Exception
                Dim x = MsgBox("This Record is not available to the Sales Department." & vbCr & vbCr & "Would you like to make it available add it to ""Unfiltered Sales Dept. List""?", MsgBoxStyle.YesNo, "Cannot Access Record")
                Me.lvSales.Focus()
                If x = vbYes Then

                    Dim cnn1 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
                    Dim cmdGet1 As SqlCommand
                    Dim r1 As SqlDataReader
                    cmdGet1 = New SqlCommand("Update tblwherecanleadgo set Sales = 'True' Where Leadnumber = " & CType(Me.txtSingleRecordInput.Text, Integer), cnn1)
                    cmdGet1.CommandType = CommandType.Text

                    cnn1.Open()
                    r1 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)

                    r1.Read()

                    r1.Close()
                    cnn1.Close()
                    Dim c As New SalesListManager
                    Me.lvSales.Items(Me.txtSingleRecordInput.Text).Selected = True
                    Me.lvSales.SelectedItems(0).EnsureVisible()
                End If
            End Try
        End Try


        Me.txtSingleRecordInput.Text = "[Enter Record ID]"
        Me.btnSingleRecord.Enabled = False
        Me.btnBuildList.HideDropDown()
    End Sub


    Private Sub btnSaveRep_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveRep.Click
        Dim cnn2 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet1 As SqlCommand
        Dim r2 As SqlDataReader
        cmdGet1 = New SqlCommand("Update enterlead set rep1 = " & "'" & Me.cboRep1.Text & "'" & ", rep2 = " & "'" & Me.cboRep2.Text & "'" & ",lastupdated = 'true' where ID = " & Me.ID, cnn2)
        cmdGet1.CommandType = CommandType.Text
        cnn2.Open()
        r2 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
        r2.Read()
        '' Loads Names of Latest Rep from sales rep pull list and will also add 
        '' the name to rep combos if they are not part of the current rep list 
        Try
            Me.cboRep1.Text = r2.Item(0)
            If Me.cboRep1.Text = "" And r2.Item(0) <> "" Then
                Me.cboRep1.Items.Add(r2.Item(0))
                Me.cboRep2.Items.Add(r2.Item(0))
                Me.cboRep1.Text = r2.Item(0)

            End If
        Catch ex As Exception
            Me.cboRep1.Text = Nothing

        End Try
        Try
            Me.cboRep2.Text = r2.Item(1)
            If Me.cboRep2.Text = "" And r2.Item(1) <> "" Then
                Me.cboRep2.Items.Add(r2.Item(1))
                Me.cboRep1.Items.Add(r2.Item(1))
                Me.cboRep2.Text = r2.Item(1)

            End If
        Catch ex As Exception
            Me.cboRep2.Text = Nothing

        End Try

        r2.Close()
        cnn2.Close()
        If Me.TabControl2.SelectedIndex = 1 Then
            Me.lvMemorized_SelectedIndexChanged(Nothing, Nothing)
        Else
            Me.lvSales_SelectedIndexChanged(Nothing, Nothing)
        End If



    End Sub

    Private Sub btnCustomerTools_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCustomerTools.Click
        If Me.ID = "" Then
            Me.btnSaveRep.Enabled = False
        Else
            Me.btnSaveRep.Enabled = True
        End If
    End Sub

    Private Sub btnMemorize_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMemorize.Click
        If Me.ID = "" Then
            MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If

        If Me.btnMemorize.Text = "Memorize This Record" Then
          
            MemorizeNotes.ShowDialog()
        Else
            Dim cnn2 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdGet1 As SqlCommand
            Dim r2 As SqlDataReader
            cmdGet1 = New SqlCommand("Delete memorizedappts where leadnum = " & "'" & STATIC_VARIABLES.CurrentID & "'" & " and UserLoggedOn = " & "'" & STATIC_VARIABLES.CurrentUser & "'" & " and Form = 'Sales' ", cnn2)
            cmdGet1.CommandType = CommandType.Text
            cnn2.Open()
            r2 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
            r2.Read()
            r2.Close()
            cnn2.Close()
            Me.PopulateMemorized()


        End If
    End Sub



    Private Sub lblFilterGroups_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFilterGroups.Click
        Me.cboFilterGroups.Focus()
        Me.cboFilterGroups.DroppedDown = True
    End Sub



    Private Sub cboFilterGroups_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFilterGroups.SelectedIndexChanged

        If Me.cboFilterGroups.Text <> "" Then
            Me.lblFilterGroups.Text = Me.cboFilterGroups.Text
            Me.lblFilterGroups.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblFilterGroups.ForeColor = Color.Black
            Me.cboGroupByMemorized.Enabled = True
            Me.lblgroupbymemorized.Enabled = True
        Else
            Me.lblFilterGroups.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblFilterGroups.ForeColor = Color.Gray
            Me.lblFilterGroups.Text = "Filter Groups"
            Me.cboGroupByMemorized.SelectedItem = Nothing
            Me.cboGroupByMemorized.Enabled = False
            Me.lblgroupbymemorized.Enabled = False
        End If
        Me.PopulateMemorized()
    End Sub

 

    Private Sub cboGroupByMemorized_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupByMemorized.SelectedIndexChanged
        If Me.cboGroupByMemorized.Text <> "" Then
            Me.lblgroupbymemorized.Text = Me.cboGroupByMemorized.Text
            Me.lblgroupbymemorized.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblgroupbymemorized.ForeColor = Color.Black
        Else
            Me.lblgroupbymemorized.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblgroupbymemorized.ForeColor = Color.Gray
            Me.lblgroupbymemorized.Text = "Group By"
        End If
        Me.PopulateMemorized()
    End Sub

    Private Sub lblgroupbymemorized_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblgroupbymemorized.Click
        Me.cboGroupByMemorized.Focus()
        Me.cboGroupByMemorized.DroppedDown = True
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        If Me.TabControl2.SelectedIndex = 0 Then
            Me.ToolbarConfig(2)
            If Me.lvSales.Items.Count = 0 Or Me.lvSales.SelectedItems.Count = 0 Then
                Me.PullInfo("")
                Me.ID = ""
                STATIC_VARIABLES.CurrentID = Me.ID
            Else
                Me.PullInfo(Me.lvSales.SelectedItems(0).Text)
                Me.ID = Me.lvSales.SelectedItems(0).Text
                STATIC_VARIABLES.CurrentID = Me.ID
            End If

        Else
            Me.ToolbarConfig(3)
            Me.PopulateMemorized()
            If Me.lvMemorized.Items.Count = 0 Or Me.lvMemorized.SelectedItems.Count = 0 Then
                Me.PullInfo("")
                Me.ID = ""
                STATIC_VARIABLES.CurrentID = Me.ID
            Else
                Me.PullInfo(Me.lvMemorized.SelectedItems(0).Tag)
                Me.ID = Me.lvMemorized.SelectedItems(0).Tag
                STATIC_VARIABLES.CurrentID = Me.ID
            End If
            If Me.cboFilterGroups.SelectedItem = Nothing Or Me.cboFilterGroups.SelectedItem = "" Then
                Me.cboGroupByMemorized.Enabled = False
                Me.lblgroupbymemorized.Enabled = False
            End If
      
        End If
    End Sub

    Private Sub btnExpandMemorize_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExpandMemorize.Click

        Me.btnExpandSalesList_Click(Nothing, Nothing)

    End Sub

    Private Sub lbldisplaymemorized_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbldisplaymemorized.Click
        Me.cboDisplayMemorized.Focus()
        Me.cboDisplayMemorized.DroppedDown = True
    End Sub

    Private Sub lblDisplayColumn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblDisplayColumn.Click
        Me.cboDisplayColumn.Focus()
        Me.cboDisplayColumn.DroppedDown = True
    End Sub


 



    Private Sub cboDisplayMemorized_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDisplayMemorized.SelectedIndexChanged
        If Me.cboDisplayMemorized.Text <> "" Then
            Me.lbldisplaymemorized.Text = Me.cboDisplayMemorized.Text
            Me.lbldisplaymemorized.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lbldisplaymemorized.ForeColor = Color.Black
        Else
            Me.lbldisplaymemorized.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lbldisplaymemorized.ForeColor = Color.Gray
            Me.lbldisplaymemorized.Text = "Change Display Column"
        End If
        Me.DisplayColumn(Me.cboDisplayMemorized.Text, Me.lvMemorized)
    End Sub

    Private Sub lnkOrderbyMem_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkOrderbyMem.LinkClicked
        If Me.lnkOrderbyMem.Text = "Order By ID" Then

            Me.lnkOrderbyMem.Text = "Order By Appt. Date"
            Me.lnkOrderbyMem.Location = New System.Drawing.Point(33, 3)
            Me.PopulateMemorized()
        Else

            Me.lnkOrderbyMem.Text = "Order By ID"
            Me.lnkOrderbyMem.Location = New System.Drawing.Point(54, 3)
            Me.PopulateMemorized()
        End If
    End Sub

    Private Sub ShowNotesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowNotesToolStripMenuItem.Click
        Me.lvMemorized_DoubleClick(Nothing, Nothing)
    End Sub

    Private Sub EnterSalesResultToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnterSalesResultToolStripMenuItem.Click
        Me.btnSalesResult2_ButtonClick(Nothing, Nothing)
    End Sub

    Private Sub MemorizeThisApptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MemorizeThisApptToolStripMenuItem.Click
        Me.btnMemorize_Click(Nothing, Nothing)
    End Sub

    Private Sub RemoveThisApptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveThisApptToolStripMenuItem.Click
        Me.btnMemorize_Click(Nothing, Nothing)
    End Sub

    Private Sub CustomerToolsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerToolsToolStripMenuItem.Click

    End Sub

    Private Sub SetAppointmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SetAppointmentToolStripMenuItem.Click
        Me.btnSetAppt_Click(Nothing, Nothing)
    End Sub

    Private Sub SaveChangesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveChangesToolStripMenuItem.Click
        Me.btnSaveRep_Click(Nothing, Nothing)
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Me.cboRep1.Text = Me.ToolStripComboBox1.Text
    End Sub

    Private Sub ToolStripComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged
        Me.cboRep2.Text = Me.ToolStripComboBox2.Text
    End Sub


 

  

    Private Sub btnEditCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEditCustomer.Click
        If Me.TabControl2.SelectedIndex = 0 Then
            If Me.lvSales.SelectedItems.Count <> 0 Then
                EditCustomerInfo.ID = STATIC_VARIABLES.CurrentID
            End If
        Else
            If Me.lvMemorized.SelectedItems.Count <> 0 Then
                EditCustomerInfo.ID = STATIC_VARIABLES.CurrentID
            End If
        End If
        EditCustomerInfo.Show()

    End Sub



    Private Sub MarkTaskAsDoneToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MarkTaskAsDoneToolStripMenuItem.Click
        Me.btnMarkTaskAsDone_Click(Nothing, Nothing)
    End Sub

    Private Sub EditScheduledTaskToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditScheduledTaskToolStripMenuItem.Click
        Me.btnEditScheduledTask_Click(Nothing, Nothing)
    End Sub

    Private Sub btnEditScheduledTask_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEditScheduledTask.Click
        ScheduleAction.edit = True
        Dim c As Integer = Me.pnlScheduledTasks.Controls.Count
        Dim i As Integer
        For i = 1 To c
            Dim all As Panel = Me.pnlScheduledTasks.Controls(i - 1)
            If all.BorderStyle = BorderStyle.FixedSingle Then
                ScheduleAction.EditId = all.Name.ToString.Substring(3)
            End If
        Next
        If ScheduleAction.EditId = "" Then
            MsgBox("You must select a Task.", MsgBoxStyle.Exclamation, "Please Select a Task")
            Exit Sub
        End If
        ScheduleAction.ShowDialog()
        Dim x As New ScheduledActions
        x.SetUp(Me)

    End Sub

    Private Sub HideThisCompletedTaskToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles HideThisCompletedTaskToolStripMenuItem.Click
        Me.btnRemoveThisCompletedTask_Click(Nothing, Nothing)
    End Sub

   

  
    Private Sub pnlScheduledTasks_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pnlScheduledTasks.MouseDown
        Me.HideThisCompletedTaskToolStripMenuItem.Visible = False
        Dim c As Integer = Me.pnlScheduledTasks.Controls.Count
        Dim i As Integer
        For i = 1 To c
            Dim all As Panel = Me.pnlScheduledTasks.Controls(i - 1)

            all.BorderStyle = BorderStyle.None

        Next
    End Sub

    Private Sub tsSalesDepartment_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsSalesDepartment.SizeChanged
        Me.dtpSummary.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 108, 2)
        Me.dtpSummary2.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 247, 2)
        Me.dtp1CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 247, 2)
        Me.dtp2CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 108, 2)
        Me.txt1CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 246, 4)
        Me.txt2CustomerList.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 107, 4)
        Me.dtpIssueLeads.Location = New System.Drawing.Point(Me.tsSalesDepartment.Width - 108, 2)

        ''comeback 
        '' add other datepickers and cover labels as needed 
    End Sub

    Private Sub SCsalesresult_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SCsalesresult.Click
        Me.btnSalesResult_Click1(Nothing, Nothing)
    End Sub

    Private Sub SCCustomerList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SCCustomerList.Click
        Me.lvnoresults_DoubleClick(Nothing, Nothing)
    End Sub

    Private Sub cboDateRangeSummary_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDateRangeSummary.SelectedIndexChanged
        If Me.LoadComplete = False Then
            Exit Sub
        End If
     

        If Me.cboDateRangeSummary.Text <> "Custom" Then
            Dim d As New DTPManipulation(Me.cboDateRangeSummary.Text)
            Me.dtpSummary2.Value = d.retDateFrom
            Me.dtpSummary.Value = d.retDateTo

        End If
        Dim r As New Sales_Performance_Report

    End Sub
    Dim dtpsum2orig As String
    Dim dtpsum1orig As String
    Dim Focusdtp1 As Boolean = False
    Dim focusdtp2 As Boolean = False

    Private Sub dtpSummary_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpSummary.GotFocus
        If LoadComplete = False Then
            Exit Sub
        End If
        Focusdtp1 = True
        dtpsum1orig = Me.dtpSummary.Value.ToString
     


    End Sub
 

    Private Sub dtpSummary2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpSummary2.GotFocus
        If LoadComplete = False Then
            Exit Sub
        End If
        focusdtp2 = True
        dtpsum2orig = Me.dtpSummary2.Value.ToString
     
    End Sub

    Private Sub lvAttachedFiles_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        'If Me.lvAttachedFiles.SelectedItems.Count <> 0 Then
        '    Dim x As New jbPicsAttachedFiles
        '    x.LaunchAttachedFile(Me.ID, Me.lvAttachedFiles.SelectedItems.Item(0).Text)
        'End If
    End Sub

    Private Sub lvAttachedFiles_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub lvAttachedFiles_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

    End Sub

    Private Sub lvAttachedFiles_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'If Me.lvAttachedFiles.SelectedItems.Count <> 0 Then
        '    Me.cmsAF.Items(0).Visible = True
        '    Me.cmsAF.Items(1).Visible = True
        '    Me.cmsAF.Items(2).Visible = True
        '    Me.cmsAF.Items(3).Visible = True
        '    Me.cmsAF.Items(4).Visible = True
        '    Me.cmsAF.Items(5).Visible = True
        '    Me.cmsAF.Items(6).Visible = True
        '    Me.cmsAF.Items(7).Visible = True
        '    Me.cmsAF.Items(8).Visible = True
        '    Me.cmsAF.Items(9).Visible = True
        '    Me.cmsAF.Items(10).Visible = True
        '    Me.cmsAF.Items(11).Visible = False
        '    Me.cmsAF.Items(12).Visible = False
        '    Me.cmsAF.Items(13).Visible = False
        '    Me.cmsAF.Items(14).Visible = False
        '    Me.cmsAF.Items(15).Visible = False
        '    Me.cmsAF.Items(16).Visible = False
        '    Me.cmsAF.Items(17).Visible = False
        'Else
        '    Me.cmsAF.Items(0).Visible = False
        '    Me.cmsAF.Items(1).Visible = False
        '    Me.cmsAF.Items(2).Visible = False
        '    Me.cmsAF.Items(3).Visible = False
        '    Me.cmsAF.Items(4).Visible = False
        '    Me.cmsAF.Items(5).Visible = False
        '    Me.cmsAF.Items(6).Visible = False
        '    Me.cmsAF.Items(7).Visible = False
        '    Me.cmsAF.Items(8).Visible = False
        '    Me.cmsAF.Items(9).Visible = False
        '    Me.cmsAF.Items(10).Visible = False
        '    Me.cmsAF.Items(11).Visible = True
        '    Me.cmsAF.Items(12).Visible = True
        '    Me.cmsAF.Items(13).Visible = True
        '    Me.cmsAF.Items(14).Visible = True
        '    Me.cmsAF.Items(15).Visible = True
        '    Me.cmsAF.Items(16).Visible = True
        '    Me.cmsAF.Items(17).Visible = True

        'End If
    End Sub

 
   
    Private Sub lvAttachedFiles_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Dim g As New jbPicsAttachedFiles


    Private Sub btnOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpen.Click
        'Dim x As New jbPicsAttachedFiles
        'x.LaunchAttachedFile(Me.ID, Me.lvAttachedFiles.SelectedItems.Item(0).Text)
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        'Dim b As ListViewItem
        'For Each b In Me.lvAttachedFiles.Items
        '    If b.Selected = True Then
        '        '' this is the one to delete 
        '        '' 
        '        g.DeleteFile(b.Text)
        '    End If
        'Next
        'Dim z As New PopulateAF(Me)
    End Sub

    Private Sub btnRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRename.Click
        'If Me.lvAttachedFiles.SelectedItems.Count > 1 Then
        '    MsgBox("You can only rename 1 file at a time!", MsgBoxStyle.Exclamation, "Rename does not support multiple files")
        '    Exit Sub
        'End If


        'Dim y As String = InputBox("New File Name", "New File Name")
        'If y = "" Then
        '    Exit Sub
        'End If
        'Dim count As Integer = 0

        'If y.Contains("/") Then
        '    count += 1
        'ElseIf y.Contains("\") Then
        '    count += 1
        'ElseIf y.Contains("?") Then
        '    count += 1
        'ElseIf y.Contains(":") Then
        '    count += 1
        'ElseIf y.Contains("*") Then
        '    count += 1
        'ElseIf y.Contains(">") Then
        '    count += 1
        'ElseIf y.Contains("<") Then
        '    count += 1
        'ElseIf y.Contains("|") Then
        '    count += 1
        'ElseIf y.Contains(Chr(34)) Then
        '    count += 1
        'End If
        'If count >= 1 Then
        '    MsgBox("Cannot use these Characters: \ / ? : *  > < | " & Chr(34), MsgBoxStyle.Exclamation, "Cannot use Special Characters in File Name")
        '    Exit Sub
        'End If
        'Dim s = Split(Me.lvAttachedFiles.SelectedItems.Item(0).Text, ".")
        'Dim x As New jbPicsAttachedFiles
        'x.RenameFile(s(0), y, s(1))
        'Dim z As New PopulateAF(Me)
    End Sub

    Private Sub btnCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCut.Click
        'Dim h As ListViewItem
        'g._arCopiedAttachedFiles.Clear()
        'g._arCopiedOperation.Clear()
        'Clipboard.Clear()
        'For Each h In Me.lvAttachedFiles.Items
        '    If h.Selected = True Then
        '        Dim arNameAndExt = Split(h.Text, ".")
        '        Dim FullName As String = g.ConvertFilePath(arNameAndExt(0), arNameAndExt(1))
        '        g._arCopiedAttachedFiles.Add(FullName)
        '        g._arCopiedOperation.Add("Cut")
        '        Me.lvAttachedFiles.Items.Remove(h)
        '    End If
        'Next
    End Sub

    Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        'Try

        '    Dim a As ListViewItem
        '    For Each a In Me.lvAttachedFiles.Items
        '        If a.Selected = True Then
        '            Dim arNameAndExt = Split(a.Text, ".")
        '            Dim FullName As String = g.ConvertFilePath(arNameAndExt(0), arNameAndExt(1))
        '            g._arCopiedAttachedFiles.Add(FullName)
        '            g._arCopiedOperation.Add("Copy")
        '        End If
        '    Next
        'Catch ex As Exception
        '    '' catch the error and call log 
        'End Try
    End Sub

    Private Sub btnAscending_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAscending.Click
        Me.btnAscending.Checked = True
        Me.btnDescending.Checked = False
    End Sub

    Private Sub btnDescending_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDescending.Click
        Me.btnAscending.Checked = False
        Me.btnDescending.Checked = True
    End Sub

    Private Sub btnXLarge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnXLarge.Click
        Me.btnXLarge.Checked = True
        Me.btnLarge.Checked = False
        Me.btnMedium.Checked = False
        Me.btnSmall.Checked = False
        Me.btnList.Checked = False
        Me.btnDetails.Checked = False
        Me.btnTiles.Checked = False
    End Sub

    Private Sub btnLarge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLarge.Click
        Me.btnXLarge.Checked = False
        Me.btnLarge.Checked = True
        Me.btnMedium.Checked = False
        Me.btnSmall.Checked = False
        Me.btnList.Checked = False
        Me.btnDetails.Checked = False
        Me.btnTiles.Checked = False
    End Sub

    Private Sub btnMedium_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMedium.Click
        Me.btnXLarge.Checked = False
        Me.btnLarge.Checked = False
        Me.btnMedium.Checked = True
        Me.btnSmall.Checked = False
        Me.btnList.Checked = False
        Me.btnDetails.Checked = False
        Me.btnTiles.Checked = False
    End Sub

    Private Sub btnSmall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSmall.Click
        Me.btnXLarge.Checked = False
        Me.btnLarge.Checked = False
        Me.btnMedium.Checked = False
        Me.btnSmall.Checked = True
        Me.btnList.Checked = False
        Me.btnDetails.Checked = False
        Me.btnTiles.Checked = False
    End Sub

    Private Sub btnList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnList.Click
        Me.btnXLarge.Checked = False
        Me.btnLarge.Checked = False
        Me.btnMedium.Checked = False
        Me.btnSmall.Checked = False
        Me.btnList.Checked = True
        Me.btnDetails.Checked = False
        Me.btnTiles.Checked = False
    End Sub

    Private Sub btnDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetails.Click
        Me.btnXLarge.Checked = False
        Me.btnLarge.Checked = False
        Me.btnMedium.Checked = False
        Me.btnSmall.Checked = False
        Me.btnList.Checked = False
        Me.btnDetails.Checked = True
        Me.btnTiles.Checked = False
    End Sub

    Private Sub btnTiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTiles.Click
        Me.btnXLarge.Checked = False
        Me.btnLarge.Checked = False
        Me.btnMedium.Checked = False
        Me.btnSmall.Checked = False
        Me.btnList.Checked = False
        Me.btnDetails.Checked = False
        Me.btnTiles.Checked = True
    End Sub

    Private Sub btnNewFolder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNewFolder.Click
        Dim y As String = InputBox("Please enter the name of the folder you want to create.", "New Folder Name", "New Folder")
        If y = "" Then
            Exit Sub
        End If
        Dim count As Integer = 0

        If y.Contains("/") Then
            count += 1
        ElseIf y.Contains("\") Then
            count += 1
        ElseIf y.Contains("?") Then
            count += 1
        ElseIf y.Contains(":") Then
            count += 1
        ElseIf y.Contains("*") Then
            count += 1
        ElseIf y.Contains(">") Then
            count += 1
        ElseIf y.Contains("<") Then
            count += 1
        ElseIf y.Contains("|") Then
            count += 1
        ElseIf y.Contains(Chr(34)) Then
            count += 1
        End If
        If count >= 1 Then
            MsgBox("Cannot use these Characters: \ / ? : *  > < | " & Chr(34), MsgBoxStyle.Exclamation, "Cannot use Special Characters in File Name")
            Exit Sub
        End If
        g.CreateDirectoryAttachedFiles(STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID & "\", y)
        'Dim z As New PopulateAF(Me)
       Me.GetRidOfOldAndPutNew()
    End Sub

    Private Sub EditCustomerToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditCustomerToolStripMenuItem.Click
        Me.btnEditCustomer_Click(Nothing, Nothing)
    End Sub

    Private Sub dtpIssueLeads_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpIssueLeads.ValueChanged
        If Me.LoadComplete = True Then
            Dim x As New Issue_Leads(True, "")
        End If

    End Sub

  
    Private Sub pnlIssue_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlIssue.Click
        For i As Integer = 1 To Me.pnlIssue.Controls.Count
            Dim all As Panel = Me.pnlIssue.Controls(i - 1)

            all.BorderStyle = BorderStyle.None

        Next
    End Sub

    Private Sub pnlScheduledTasks_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlScheduledTasks.Paint

    End Sub

    Private Sub pnlIssue_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlIssue.SizeChanged
        If LoadComplete = True Then
            If Me.panelsize < Me.pnlIssue.Width Then
                Dim x As New Issue_Leads(False, "grow")
            Else
                Dim x As New Issue_Leads(False, "shrink")
            End If


            Me.panelsize = Me.pnlIssue.Width

        End If

    End Sub



    Private Sub pnlPerformanceReport_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlPerformanceReport.SizeChanged

        Dim buffer As Integer = 25
        Dim x As Double = Me.tpSummary.Width - (94 + buffer)
        x = x / 8
        Dim h As Integer = 5
        Dim fs As Single = 9.75
        Dim j As Integer = 0
        Dim pix As Double = 0
        Dim y As Double = (x / 2) - 61
        'MsgBox(y.ToString)
        If y > 5 And y < 15 Then
            fs = 10.75
            h = 17
            j = 20
            pix = 0.5
        ElseIf y >= 15 Then
            fs = 11.75
            h = 29
            j = 20
            pix = 1
        Else
            pix = 0
        End If
        Dim fstyle As FontStyle


        If y < 0 Then
            Me.lblRep.Size = New Size(71, 20)
            Me.lblIssued.Size = New Size(122, 20)
            Me.lblResets.Size = New Size(122, 20)
            Me.lblDNS.Size = New Size(122, 20)
            Me.lblND.Size = New Size(122, 20)
            Me.lblRC.Size = New Size(122, 20)
            Me.lblsales.Size = New Size(122, 20)
            Me.lblNH.Size = New Size(122, 20)
            Me.lblSold.Size = New Size(122, 20)




            Me.lblIssued.Location = New System.Drawing.Point(74, Me.lblRep.Location.Y)
            Me.lblDNS.Location = New System.Drawing.Point(173, Me.lblRep.Location.Y)
            Me.lblNH.Location = New System.Drawing.Point(286, Me.lblRep.Location.Y)
            Me.lblResets.Location = New System.Drawing.Point(399, Me.lblRep.Location.Y)
            Me.lblND.Location = New System.Drawing.Point(512, Me.lblRep.Location.Y)
            Me.lblsales.Location = New System.Drawing.Point(625, Me.lblRep.Location.Y)
            Me.lblRC.Location = New System.Drawing.Point(738, Me.lblRep.Location.Y)
            Me.lblSold.Location = New System.Drawing.Point(851, Me.lblRep.Location.Y)
            Me.lblIssued.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblDNS.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblNH.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblResets.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblND.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblsales.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblRC.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblSold.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblRep.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            'If Me.pnlPerformanceReport.Controls.Count <= 1 Then
            '    Exit Sub
            'End If



            Try


                For q As Integer = 0 To Me.pnlPerformanceReport.Controls.Count - 1

                    Dim z
                    If Me.pnlPerformanceReport.Controls.Item(q).Text.Contains("No Data to Report") Then
                        Me.pnlPerformanceReport.Controls.Item(q).Location = New System.Drawing.Point((Me.pnlPerformanceReport.Width / 2) - 308, (Me.pnlPerformanceReport.Height / 2) - 12)

                    End If



                    If Me.pnlPerformanceReport.Controls.Item(q).Name.Contains("pnl") Then

                        z = Me.pnlPerformanceReport.Controls.Item(q)
                        If z.name = "pnl1" Then
                            h = h + 4
                            fstyle = FontStyle.Bold
                        Else
                            fstyle = FontStyle.Regular
                        End If
                        If z.name = "pnl2" Then
                            h = h - 4
                        End If

                        Dim lblissue As Label = z.Controls.Item(1)
                        Dim name As Label = z.Controls.Item(0)
                        name.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        name.Location = New System.Drawing.Point(name.Location.X, 8)

                        lblissue.Location = New System.Drawing.Point((Me.lblIssued.Location.X + 61) - (22 + h), 0)
                        lblissue.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblissue.Size = New Size(44 + h, 32)

                        Dim lblDNS As Label = z.Controls.Item(8)
                        lblDNS.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblDNS.Location = New System.Drawing.Point((Me.lblDNS.Location.X + 61) - (42 + h), 0) '192
                        lblDNS.Size = New Size(32 + h, 32)

                        Dim lblDNSper As Label = z.Controls.Item(9)
                        lblDNSper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblDNSper.Location = New System.Drawing.Point(lblDNS.Location.X + (32 + h), 0) '224
                        lblDNSper.Size = New Size(42 + h, 32)

                        Dim lblNH As Label = z.Controls.Item(4)
                        lblNH.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblNH.Size = New Size(32 + h, 32)
                        lblNH.Location = New System.Drawing.Point((Me.lblNH.Location.X + 61) - (42 + h), 0) '305

                        Dim lblNHper As Label = z.Controls.Item(5)
                        lblNHper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblNHper.Size = New Size(46 + h, 32)
                        lblNHper.Location = New System.Drawing.Point(lblNH.Location.X + (32 + h), 0) '337

                        Dim lblResets As Label = z.Controls.Item(2)
                        lblResets.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblResets.Size = New Size(32 + h, 32)
                        lblResets.Location = New System.Drawing.Point((Me.lblResets.Location.X + 61) - (42 + h), 0) '418

                        Dim lblResetsper As Label = z.Controls.Item(3)
                        lblResetsper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblResetsper.Size = New Size(46 + h, 32)
                        lblResetsper.Location = New System.Drawing.Point(lblResets.Location.X + (32 + h), 0) '450

                        Dim lblND As Label = z.Controls.Item(6)
                        lblND.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblND.Size = New Size(32 + h, 32)
                        lblND.Location = New System.Drawing.Point((Me.lblND.Location.X + 61) - (42 + h), 0) '534

                        Dim lblNDper As Label = z.Controls.Item(7)
                        lblNDper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblNDper.Size = New Size(46 + h, 32)
                        lblNDper.Location = New System.Drawing.Point(lblND.Location.X + (32 + h), 0) '566

                        Dim lblSales As Label = z.Controls.Item(12)
                        lblSales.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblSales.Size = New Size(32 + h, 32)
                        lblSales.Location = New System.Drawing.Point((Me.lblsales.Location.X + 61) - (42 + h), 0) '664

                        Dim lblSalesper As Label = z.Controls.Item(13)
                        lblSalesper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblSalesper.Size = New Size(46 + h, 32)
                        lblSalesper.Location = New System.Drawing.Point(lblSales.Location.X + (32 + h), 0) '676

                        Dim lblRC As Label = z.Controls.Item(10)
                        lblRC.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblRC.Size = New Size(32 + h, 32)
                        lblRC.Location = New System.Drawing.Point((Me.lblRC.Location.X + 61) - (42 + h), 0) '757

                        Dim lblRCper As Label = z.Controls.Item(11)
                        lblRCper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblRCper.Size = New Size(46 + h, 32)
                        lblRCper.Location = New System.Drawing.Point(lblRC.Location.X + (32 + h), 0) '789

                        Dim lblsold As Label = z.Controls.Item(14)
                        lblsold.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblsold.Size = New Size(80 + j, 32)
                        If y < 0 Then
                            lblsold.Location = New System.Drawing.Point(Me.lblSold.Location.X + 20, 0) '869
                        Else
                            lblsold.Location = New System.Drawing.Point(870, 0) '869 
                        End If



                    End If





                Next
            Catch ex As Exception
                ' MsgBox("Fail If True")
            End Try

        Else

            Dim count As Integer = 0
            Me.lblRep.Size = New Size(x, 20)
            Me.lblIssued.Size = New Size(x, 20)
            Me.lblResets.Size = New Size(x, 20)
            Me.lblDNS.Size = New Size(x, 20)
            Me.lblND.Size = New Size(x, 20)
            Me.lblRC.Size = New Size(x + 20, 20)
            Me.lblsales.Size = New Size(x, 20)
            Me.lblNH.Size = New Size(x, 20)
            Me.lblSold.Size = New Size(x, 20)

            Me.lblIssued.Location = New System.Drawing.Point(x - y, Me.lblRep.Location.Y)
            Me.lblDNS.Location = New System.Drawing.Point((x * 2) - y, Me.lblRep.Location.Y)
            Me.lblDNS.Size = New Size(Me.lblIssued.Width + j, Me.lblIssued.Height)
            Me.lblNH.Location = New System.Drawing.Point((x * 3) - y, Me.lblRep.Location.Y)
            Me.lblResets.Location = New System.Drawing.Point((x * 4) - y, Me.lblRep.Location.Y)
            Me.lblND.Location = New System.Drawing.Point((x * 5) - y, Me.lblRep.Location.Y)
            Me.lblsales.Location = New System.Drawing.Point((x * 6) - y, Me.lblRep.Location.Y)
            Me.lblRC.Location = New System.Drawing.Point((x * 7) - y, Me.lblRep.Location.Y)
            Me.lblSold.Location = New System.Drawing.Point((x * 8) - y, Me.lblRep.Location.Y)
            Me.lblIssued.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblDNS.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblNH.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblResets.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblND.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblsales.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblRC.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblSold.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblRep.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            'If Me.pnlPerformanceReport.Controls.Count <= 1 Then
            '    Exit Sub
            'End If


            Try


                For q As Integer = 0 To Me.pnlPerformanceReport.Controls.Count - 1
                    Dim z
                    If Me.pnlPerformanceReport.Controls.Item(q).Text.Contains("No Data to Report") Then
                        Me.pnlPerformanceReport.Controls.Item(q).Location = New System.Drawing.Point((Me.pnlPerformanceReport.Width / 2) - 308, (Me.pnlPerformanceReport.Height / 2) - 12)

                    End If


                    If Me.pnlPerformanceReport.Controls.Item(q).Name.Contains("pnl") Then
                        z = Me.pnlPerformanceReport.Controls.Item(q)
                        If z.name = "pnl1" Then
                            fstyle = FontStyle.Bold
                        Else
                            fstyle = FontStyle.Regular
                        End If

                        Dim name As Label = z.Controls.Item(0)

                        name.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        If name.Location.Y >= 8 Then
                            name.Location = New System.Drawing.Point(name.Location.X, name.Location.Y - pix)
                        End If


                        Dim lblissue As Label = z.Controls.Item(1)
                        lblissue.Location = New System.Drawing.Point((Me.lblIssued.Location.X + (Me.lblIssued.Width / 2)) - (22 + h), 0)
                        lblissue.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblissue.Size = New Size(44 + h, 32)


                        Dim lblDNS As Label = z.Controls.Item(8)
                        lblDNS.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblDNS.Location = New System.Drawing.Point((Me.lblDNS.Location.X + (Me.lblDNS.Width / 2)) - (42 + h), 0) '192
                        lblDNS.Size = New Size(32 + h, 32)

                        Dim lblDNSper As Label = z.Controls.Item(9)
                        lblDNSper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblDNSper.Location = New System.Drawing.Point(lblDNS.Location.X + (32 + h), 0) '224
                        lblDNSper.Size = New Size(42 + h, 32)

                        Dim lblNH As Label = z.Controls.Item(4)
                        lblNH.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblNH.Size = New Size(32 + h, 32)
                        lblNH.Location = New System.Drawing.Point((Me.lblNH.Location.X + (Me.lblNH.Width / 2)) - (42 + h), 0) '305

                        Dim lblNHper As Label = z.Controls.Item(5)
                        lblNHper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblNHper.Size = New Size(46 + h, 32)
                        lblNHper.Location = New System.Drawing.Point(lblNH.Location.X + (32 + h), 0) '337

                        Dim lblResets As Label = z.Controls.Item(2)
                        lblResets.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblResets.Size = New Size(32 + h, 32)
                        lblResets.Location = New System.Drawing.Point((Me.lblResets.Location.X + (Me.lblResets.Width / 2)) - (42 + h), 0) '418

                        Dim lblResetsper As Label = z.Controls.Item(3)
                        lblResetsper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblResetsper.Size = New Size(46 + h, 32)
                        lblResetsper.Location = New System.Drawing.Point(lblResets.Location.X + (32 + h), 0) '450

                        Dim lblND As Label = z.Controls.Item(6)
                        lblND.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblND.Size = New Size(32 + h, 32)
                        lblND.Location = New System.Drawing.Point((Me.lblND.Location.X + (Me.lblND.Width / 2)) - (42 + h), 0) '534

                        Dim lblNDper As Label = z.Controls.Item(7)
                        lblNDper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblNDper.Size = New Size(46 + h, 32)
                        lblNDper.Location = New System.Drawing.Point(lblND.Location.X + (32 + h), 0) '566

                        Dim lblSales As Label = z.Controls.Item(12)
                        lblSales.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblSales.Size = New Size(32 + h, 32)
                        lblSales.Location = New System.Drawing.Point((Me.lblsales.Location.X + (Me.lblsales.Width / 2)) - (42 + h), 0) '664

                        Dim lblSalesper As Label = z.Controls.Item(13)
                        lblSalesper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblSalesper.Size = New Size(46 + h, 32)
                        lblSalesper.Location = New System.Drawing.Point(lblSales.Location.X + (32 + h), 0) '676

                        Dim lblRC As Label = z.Controls.Item(10)
                        lblRC.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblRC.Size = New Size(32 + h, 32)
                        lblRC.Location = New System.Drawing.Point((Me.lblRC.Location.X + (Me.lblRC.Width / 2)) - (42 + h), 0) '757

                        Dim lblRCper As Label = z.Controls.Item(11)
                        lblRCper.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblRCper.Size = New Size(46 + h, 32)
                        lblRCper.Location = New System.Drawing.Point(lblRC.Location.X + (32 + h), 0) '789

                        Dim lblsold As Label = z.Controls.Item(14)
                        lblsold.Font = New System.Drawing.Font("Tahoma", fs!, fstyle, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblsold.Size = New Size(80 + j, 32)
                        If j > 0 Then
                            lblsold.Location = New System.Drawing.Point(Me.pnlPerformanceReport.Width - (lblsold.Width + 10), 0) '869
                        Else
                            lblsold.Location = New System.Drawing.Point(Me.pnlPerformanceReport.Width - (lblsold.Width + 10), 0) '869 
                        End If


                    End If
                Next

            Catch ex As Exception
                ' MsgBox("Fail Else")
            End Try
        End If

    





    End Sub

 
    Private Sub btnExclude_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExclude.Click
        If Me.btnExclude.Text.Contains("Off") Then
            Me.btnExclude.Text = "Turn On Exclusions"
            Me.btnExclude.Image = Me.ilToolbarButtons.Images(12)
        Else
            Me.btnExclude.Text = "Turn Off Exclusions"
            Me.btnExclude.Image = Me.ilToolbarButtons.Images(13)
        End If
    End Sub
    Private Sub Refocus_IssueLeads()
        Dim c As Integer = Me.pnlIssue.Controls.Count
        Dim i As Integer
        For i = 1 To c
            Dim all As Panel = Me.pnlIssue.Controls(i - 1)
            If all.BorderStyle = BorderStyle.FixedSingle Then
                STATIC_VARIABLES.CurrentID = all.Controls.Item(2).Text
            End If
        Next
    End Sub



    Private Sub btnEditCustIssue_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEditCustIssue.Click
        Refocus_IssueLeads()

        EditCustomerInfo.Show()
    End Sub

    Private Sub btnCCIssue_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCCIssue.Click
        '' edits for 'toggle button' 
        '' 8-26-2015
        ''

        '' basically a split switch 
        '' if the button text is 'this' do this
        '' elseif the button text is 'that' do that
        '' 

        Dim myBtn As ToolStripButton = sender
        Dim buttonText As String = myBtn.Text
        Dim cnt As Integer = 0
        If buttonText = "Undo Called and Cancelled" Then

            '' get the lead number
            Dim leadNum As String = ""
            Dim y As Panel
            For Each y In pnlIssue.Controls
                If y.BorderStyle = BorderStyle.FixedSingle Then
                    cnt += 1
                    Dim t As Control
                    For Each t In y.Controls
                        If TypeOf t Is LinkLabel Then
                            leadNum = t.Text

                        End If
                    Next
                End If
            Next

            If cnt <= 0 Then
                MsgBox("You must select a record to ' Undo Called and Cancelled ' . ", MsgBoxStyle.Critical, "Select a record")
                Exit Sub
            ElseIf cnt >= 1 Then
                ' MsgBox("LeadNumber : " & leadNum, MsgBoxStyle.Information, "DEBUG INFO")
                Dim b As New ToggleUndoCandC
                Dim rec_ID As String = b.Get_ID_OF_CandC(leadNum)
                b.Delete_CandC(rec_ID, leadNum, Date.Now.ToString)
                Dim c = New Issue_Leads(True, "")
            End If



        ElseIf buttonText = "Log Appt. as Called and Cancelled" Then

            Dim lead As String = ""
            For t As Integer = 0 To Me.pnlIssue.Controls.Count - 1
                If Me.pnlIssue.Controls.Item(t).Name.Contains("pnl") Then
                    Dim z As Panel = Me.pnlIssue.Controls.Item(t)
                    If z.BorderStyle = BorderStyle.FixedSingle Then
                        lead = z.Controls.Item(2).Text
                    End If
                End If
            Next
            If lead = "" Then
                MsgBox("You must Select a Record!", MsgBoxStyle.Exclamation, "No Record Selected")
                Exit Sub
            End If
            'Me.cboSalesList.SelectedItem = "Issue Leads List"
            'For x As Integer = 0 To Me.lvSales.Items.Count - 1
            '    If Me.lvSales.Items(x).Text = lead Then
            '        Me.lvSales.Items(x).Selected = True
            '        If Me.lvSales.Items(x).Selected = True Then
            '            Me.PullInfo(lead)
            '        End If
            '    End If
            'Next
            Me.PullInfo(lead)

            Dim s = Split(Me.txtContact1.Text, " ")
            Dim c1 = s(0)
            Dim s2 = Split(Me.txtContact2.Text, " ")
            Dim c2 = s2(0)
            CandCNotes.ID = lead
            CandCNotes.Contact1 = c1
            CandCNotes.Contact2 = c2
            CandCNotes.OrigApptDate = Me.txtApptDate.Text
            CandCNotes.OrigApptTime = Me.txtApptTime.Text
            CandCNotes.frm = Me


            CandCNotes.ShowInTaskbar = False
            CandCNotes.ShowDialog()

        End If
    End Sub

    Private Sub PrintThisLeadToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PrintThisLeadToolStripMenuItem.Click
        btnPrintThisIssue_Click(sender, e)
    End Sub

  
   

    Private Sub EmailThisLeadToAssignedRepsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EmailThisLeadToAssignedRepsToolStripMenuItem.Click
        btnEmailThisIssue_Click(sender, e)
    End Sub

    Private Sub btnExcludeManage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcludeManage.Click
        exclusions.ShowDialog()
    End Sub

  

    



    Private Sub Sales_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'Me.Label4.Location = New System.Drawing.Point((Me.tpSummary.Width / 2) - 87, Me.Label4.Location.Y)
        If Me.lvSales.SelectedItems.Count <> 0 Then
            Dim c As New CustomerHistory
            c.SetUp(Me, Me.lvSales.SelectedItems(0).Text, Me.TScboCustomerHistory)
        End If
        'Dim x As New ScheduledActions
        'x.SetUp("Sales")
        'Me.SplitContainer1.SplitterDistance = 2500
        Me.SplitContainer1.SplitterDistance = 218
        Me.SplitContainer1.SplitterWidth = 1
        Me.btnExpandSalesList.Text = Chr(187)
        Me.btnExpandMemorize.Text = Chr(187)


        '' EDIT: 
        '' 9-13-2015 
        '' work around for resize of control
        '' on size changed event
        '' 

        ''EDIT:
        '' 9-16-2015
        '' Need to redraw respective container on size changed
        '' If job pictures is active filter, only redraw job pictures
        '' If Attached Files is active filter......
        '' 

        Dim active_filter_text As String = Me.tscboAFPicsFilter.Text
        Select Case active_filter_text
            Case "All"
                Try
                    Dim a As Control
                    Dim p_width As Integer = Me.pnlAFPics.Width
                    Dim p_height As Integer = Me.pnlAFPics.Height
                    'x.Size = New System.Drawing.Point((p_width / 2) - 20, (p_height - 10))
                    'x.Location = New Point((p_width / 2) + 5, 0)
                    For Each a In Me.pnlAFPics.Controls

                        If TypeOf (a) Is ListView Then
                            If a.Name.ToString = "lsAF" Then
                                Dim pLocX As Integer = a.Location.X
                                Dim pLocY As Integer = a.Location.Y
                                a.Width = (p_width / 2) - 20
                                a.Height = p_height - 10
                                a.Location = New Point(5, 5)

                            ElseIf a.Name.ToString = "lsJP" Then
                                '' for list view lsJP the offset of X has to be counted in as well
                                '' 
                                Dim pLocX As Integer = (p_width / 2)
                                Dim pLocY As Integer = a.Location.Y
                                a.Width = (p_width / 2) - 20
                                a.Height = p_height - 10
                                a.Location = New Point(pLocX + 5, 5)
                            End If
                        End If
                    Next
                Catch ex As Exception
                    '' just trap it

                End Try
                Exit Select
            Case "Job Pictures"
                Try
                    Dim a As Control
                    Dim p_width As Integer = Me.pnlAFPics.Width
                    Dim p_height As Integer = Me.pnlAFPics.Height

                    For Each a In Me.pnlAFPics.Controls

                        If TypeOf (a) Is ListView Then

                            If a.Name.ToString = "lsJP" Then
                                Dim pLocX As Integer = (p_width)
                                Dim pLocY As Integer = a.Location.Y
                                a.Width = (p_width - 20)
                                a.Height = (p_height - 10)
                                a.Location = New Point(5, 5)
                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try
                Exit Select
            Case "Attached Files"
                Try
                    Dim a As Control
                    Dim p_width As Integer = Me.pnlAFPics.Width
                    Dim p_height As Integer = Me.pnlAFPics.Height

                    For Each a In Me.pnlAFPics.Controls

                        If TypeOf (a) Is ListView Then

                            If a.Name.ToString = "lsAF" Then
                                Dim pLocX As Integer = a.Location.X
                                Dim pLocY As Integer = a.Location.Y
                                a.Width = (p_width) - 20
                                a.Height = p_height - 10
                                a.Location = New Point(5, 5)

                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try
                Exit Select
            Case Else
                Try
                    Dim a As Control
                    Dim p_width As Integer = Me.pnlAFPics.Width
                    Dim p_height As Integer = Me.pnlAFPics.Height
                    'x.Size = New System.Drawing.Point((p_width / 2) - 20, (p_height - 10))
                    'x.Location = New Point((p_width / 2) + 5, 0)
                    For Each a In Me.pnlAFPics.Controls

                        If TypeOf (a) Is ListView Then
                            If a.Name.ToString = "lsAF" Then
                                Dim pLocX As Integer = a.Location.X
                                Dim pLocY As Integer = a.Location.Y
                                a.Width = (p_width / 2) - 20
                                a.Height = p_height - 10
                                a.Location = New Point(5, 5)

                            ElseIf a.Name.ToString = "lsJP" Then
                                '' for list view lsJP the offset of X has to be counted in as well
                                '' 
                                Dim pLocX As Integer = (p_width / 2)
                                Dim pLocY As Integer = a.Location.Y
                                a.Width = (p_width / 2) - 20
                                a.Height = p_height - 10
                                a.Location = New Point(pLocX + 5, 5)
                            End If
                        End If
                    Next
                Catch ex As Exception
                    '' just trap it

                End Try
                Exit Select
        End Select

    End Sub

    Private Sub tpSummary_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tpSummary.SizeChanged
        Me.Label4.Location = New System.Drawing.Point((Me.tpSummary.Width / 2) - 87, Me.Label4.Location.Y)
    End Sub




    Private Sub btnSetAppt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetAppt.Click
        If Me.ID = "" Then
            MsgBox("You must select a record to Set Appointment!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        SetAppt.ID = Me.ID
        SetAppt.frm = Me

        SetAppt.OrigApptDate = Me.txtApptDate.Text
        SetAppt.OrigApptTime = Me.txtApptTime.Text
        Dim s = Split(Me.txtContact1.Text, " ")
        Dim s2 = Split(Me.txtContact2.Text, " ")
        SetAppt.Contact1 = s(0)
        SetAppt.Contact2 = s2(0)
        If STATIC_VARIABLES.salesworkaround = True Then
            STATIC_VARIABLES.salesworkaround = False
            SetAppt.ShowDialog()
            Me.btnSetAppt_Click(sender, e)
        Else
            SetAppt.ShowDialog()
        End If

    
    End Sub

#Region " ' Email ' buttons tracked down from tool strip sales department . . .  "
    ''
    '' 8-15-2015
    '' 3 nested ' buttons ' under btnEmailThisIssue
    '' 


    Private Sub btnEmailIssue_Click(sender As Object, e As EventArgs) Handles btnEmailIssue.Click
        'MsgBox("test hit btnEmailIssue", MsgBoxStyle.Information, "DEBUG INFO")
    End Sub

    Private Sub btnEmailIssuedAppts_Click(sender As Object, e As EventArgs) Handles btnEmailIssuedAppts.Click
        'MsgBox("test hit btnEmailIssuedAppts - marketing manager button")
        Dim xyz As New EmailIssuedLeads
        Dim arMSGS As New ArrayList
        arMSGS = xyz.Generate_MarketingManager_List(Me.dtpIssueLeads.Value.ToString)
        Dim mmOBJ As EmailIssuedLeads.MarketingManagerOBJ
        xyz.MailTheListToMarketingManager(arMSGS, mmOBJ.FName, mmOBJ.LName, mmOBJ.EmailAddress, Me.dtpIssueLeads.Value.ToString)
        '' marketing manager list
    End Sub

    Private Sub btnEmailAllIssue_Click(sender As Object, e As EventArgs) Handles btnEmailAllIssue.Click
        ' MsgBox("test hit btnEmailallIssue ")
        '' email to all reps that can get
        Dim z As New EmailIssuedLeads
        Dim arRepsThatCanGetEmail As New ArrayList
        Dim arRepsThatDontGetEmail As New ArrayList
        Dim arLeadNum As New ArrayList
        Dim emlBody As String = ""
        Dim iteration As Integer = 0
        If Me.btnExclude.Text.Contains("Off") Then
            Dim y As Panel
            Dim leadNum As String = "0"
            For Each y In pnlIssue.Controls
                If y.BorderStyle = BorderStyle.None Then
                    '' this is the selected record now determine rep and can they get email and so on. . . . . 
                    Dim ctrl As Control
                    For Each ctrl In y.Controls
                        If TypeOf ctrl Is LinkLabel Then
                            leadNum = ctrl.Text
                            '' if its a called an cancelled, dont do anything with it. 
                            '' called and cancelled will come from enter lead
                            '' proc = GetIssuedCancels
                           
                        End If
                        If TypeOf ctrl Is ComboBox And ctrl.Text <> "" Then
                            If y.Controls.ContainsKey("lnk" & leadNum.ToString) = True Then
                                
                                Dim splitName() = Split(ctrl.Text, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                                Dim canGetEmail As Boolean = z.CanRepGetEmail(splitName(0), splitName(1))
                                If canGetEmail = True Then
                                    arRepsThatCanGetEmail.Add(splitName(0) & " " & splitName(1) & " | " & leadNum.ToString)
                                ElseIf canGetEmail = False Then
                                    arRepsThatDontGetEmail.Add(splitName(0) & " " & splitName(1) & " | " & leadNum.ToString)
                                End If
                            End If

                        End If
                    Next
                End If
            Next

            Dim iii As Integer = 0
            Dim _doo As String = ""
            For iii = 0 To arRepsThatCanGetEmail.Count - 1
                _doo += arRepsThatCanGetEmail(iii) & vbCrLf
            Next
            Dim aaa As Integer = 0
            Dim _dontt As String
            For aaa = 0 To arRepsThatDontGetEmail.Count - 1
                _dontt += arRepsThatDontGetEmail(aaa) & vbCrLf
            Next
            ' MsgBox("DO: " & vbCrLf & _doo & vbCrLf & "DONT:" & vbCrLf & _dontt, MsgBoxStyle.Information, "Debug NFO - Reps that DO: With Exclusions")
            Dim xyz As New EmailIssuedLeads
            Dim arStripLeadOut As New ArrayList

            '' now strip off called an cancelled from array

            Dim dd As Integer = 0
            Dim iterationR As Integer = 0
            For dd = 0 To arRepsThatCanGetEmail.Count - 1
                iterationR += 1
                Dim recID = Split(arRepsThatCanGetEmail(dd), " | ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                Dim id_ = recID(1)
                If xyz.GetEnterLeadMarketingResult(id_) = True Then
                    arStripLeadOut.Add(iterationR - 1)
                End If
            Next

            Dim ff As Integer = 0
            For ff = 0 To arStripLeadOut.Count - 1
                arRepsThatCanGetEmail.RemoveAt(arStripLeadOut(ff))
            Next

            '' after its all said and done, mail what is left.

            xyz.BulkMailWithExclusions(arRepsThatCanGetEmail, Me.dtpIssueLeads.Value.ToString)





        End If
        If Me.btnExclude.Text.Contains("On") Then
            Dim y As Panel
            Dim leadNum As String = "0"
            For Each y In pnlIssue.Controls
                If y.BorderStyle = BorderStyle.None Then
                    '' this is the selected record now determine rep and can they get email and so on. . . . . 
                    Dim ctrl As Control
                    For Each ctrl In y.Controls
                        If TypeOf ctrl Is LinkLabel Then
                            leadNum = ctrl.Text

                            '' if its a called an cancelled, dont do anything with it. 
                            '' called and cancelled will come from enter lead
                            '' proc = GetIssuedCancels
                            

                           
                        End If
                        If TypeOf ctrl Is ComboBox And ctrl.Text <> "" Then
                            If y.Controls.ContainsKey("lnk" & leadNum.ToString) = True Then
                                 
                                Dim splitName() = Split(ctrl.Text, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                                Dim canGetEmail As Boolean = z.CanRepGetEmail(splitName(0), splitName(1))
                                If canGetEmail = True Then
                                    arRepsThatCanGetEmail.Add(splitName(0) & " " & splitName(1) & " | " & leadNum.ToString)
                                ElseIf canGetEmail = False Then
                                    arRepsThatDontGetEmail.Add(splitName(0) & " " & splitName(1) & " | " & leadNum.ToString)
                                End If
                            End If

                        End If
                    Next
                End If
            Next
        Dim ii As Integer = 0
        Dim _do As String = ""
        For ii = 0 To arRepsThatCanGetEmail.Count - 1
            _do += arRepsThatCanGetEmail(ii) & vbCrLf
        Next
        Dim aa As Integer = 0
        Dim _dont As String
        For aa = 0 To arRepsThatDontGetEmail.Count - 1
            _dont += arRepsThatDontGetEmail(aa) & vbCrLf
            Next
            '' DEBUG INFO
            'MsgBox("DO: " & vbCrLf & _do & vbCrLf & "DONT:" & vbCrLf & _dont, MsgBoxStyle.Information, "Debug NFO - Reps That DO: No Exclusions")

            Dim xyz As New EmailIssuedLeads
            Dim arStripLeadOut As New ArrayList

            '' now strip off called an cancelled from array

            Dim dd As Integer = 0
            Dim iterationR As Integer = 0
            For dd = 0 To arRepsThatCanGetEmail.Count - 1
                iterationR += 1
                Dim recID = Split(arRepsThatCanGetEmail(dd), " | ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                Dim id_ = recID(1)
                If xyz.GetEnterLeadMarketingResult(id_) = True Then
                    arStripLeadOut.Add(iterationR - 1)
                End If
            Next

            Dim ff As Integer = 0
            For ff = 0 To arStripLeadOut.Count - 1
                arRepsThatCanGetEmail.RemoveAt(arStripLeadOut(ff))
            Next

            xyz.BulkEmailWithoutExceptions(arRepsThatCanGetEmail, Me.dtpIssueLeads.Value.ToString)
        End If
       

    End Sub

    Private Sub btnEmailThisIssue_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmailThisIssue.Click

        '' individual selected lead to be emailed 
        '' 

        EmailSingleRecord()



    End Sub

#End Region

    Private Sub EmailSingleRecord()
        Dim z As New EmailIssuedLeads
        Dim emlBody As String = ""
        If Me.btnExclude.Text.Contains("Off") Then
            Dim y As Panel
            Dim leadNum As String = "0"
            For Each y In pnlIssue.Controls
                If y.BorderStyle = BorderStyle.FixedSingle Then
                    '' this is the selected record now determine rep and can they get email and so on. . . . . 
                    Dim ctrl As Control
                    For Each ctrl In y.Controls
                        If TypeOf ctrl Is LinkLabel Then
                            'If InStr(ctrl.Name, "lnk", Microsoft.VisualBasic.CompareMethod.Text) = True Then
                            leadNum = ctrl.Text
                            '' check last m result here to make sure it can be done. 
                            '' 
                            ''--8-25-2015 Per Andy Edit
                            '' needs to filter down to ALL print options, all Email options
                            '' 
                            


                        ElseIf TypeOf ctrl Is ComboBox And ctrl.Text <> "" Then
                            Dim strRep() = Split(ctrl.Text, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                            Dim fname As String = strRep(0)
                            Dim lname As String = strRep(1)

                            Dim canGetEmail As Boolean = z.CanRepGetEmail(fname, lname)
                            If canGetEmail = False Then
                                MsgBox("This sales rep cannot recieve emails.", MsgBoxStyle.Critical, "Can't Recieve Email")
                                Exit Sub
                            End If
                            Dim emailAddy As String = z.GetRepEmailAddress(fname, lname)
                            Dim exclusionSet As EmailIssuedLeads.Exclusions = z.GetExclusions()
                            STATIC_VARIABLES.CurrentExclusionSet = exclusionSet
                            'MsgBox(fname & " | " & lname & vbCrLf & "Email ? : " & canGetEmail & vbCrLf & "Email Address: " & emailAddy & vbCrLf & "Lead Num: " & leadNum & vbCrLf & vbCrLf & "Exclusions :>" & vbCrLf & "Generated: " & exclusionSet.Generated & vbCrLf & "Marketer: " & exclusionSet.Marketer & vbCrLf & "PLS: " & exclusionSet.PLS & " | SLS: " & exclusionSet.SLS & vbCrLf & "LastMResult:" & exclusionSet.LastMResult & vbCrLf & "Phone: " & exclusionSet.Phone, MsgBoxStyle.Information, "Debug Info:")
                            If canGetEmail = True Then
                                '' check here to make sure not called an cancelled
                                '' 8-26-2014

                                If z.IsCalledAndCancelled(leadNum, Me.dtpIssueLeads.Value.ToShortDateString) = False Then
                                    MsgBox("This lead was called and cancelled." & vbCrLf & "Please check your marketing results and try again.", MsgBoxStyle.Critical, "Lead Called and Cancelled")
                                    Exit Sub
                                End If

                                emlBody = z.ConstructMessageWithExclusions(fname, lname, leadNum, exclusionSet, emailAddy)
                                'MsgBox(emlBody)
                                '' uncomment here to actually send
                                z.EMAIL_SINGLE_MarkupEmail_WITH_EXCLUSIONS(fname, lname, leadNum, exclusionSet, emailAddy, emlBody, "Record ID: " & leadNum.ToString)
                            ElseIf canGetEmail = False Then
                                '' do nothing for now. 
                                ''
                            End If

                        End If '' end type of control 

                    Next
                End If
            Next
        ElseIf Me.btnExclude.Text.Contains("On") Then
            Dim y As Panel
            Dim leadNum As String = "0"
            For Each y In pnlIssue.Controls
                If y.BorderStyle = BorderStyle.FixedSingle Then
                    '' this is the selected record now determine rep and can they get email and so on. . . . . 
                    Dim ctrl As Control
                    For Each ctrl In y.Controls
                        If TypeOf ctrl Is LinkLabel Then
                            'If InStr(ctrl.Name, "lnk", Microsoft.VisualBasic.CompareMethod.Text) = True Then
                            leadNum = ctrl.Text
                            
                        ElseIf TypeOf ctrl Is ComboBox And ctrl.Text <> "" Then
                            Dim strRep() = Split(ctrl.Text, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                            Dim fname As String = strRep(0)
                            Dim lname As String = strRep(1)
                            Dim canGetEmail As Boolean = z.CanRepGetEmail(fname, lname)
                            If canGetEmail = False Then
                                MsgBox("This Sales Rep cannot recieve Email.", MsgBoxStyle.Critical, "Can't Recieve Email")
                                Exit Sub
                            End If
                            Dim emailAddy As String = z.GetRepEmailAddress(fname, lname)
                            Dim exclusionSet As EmailIssuedLeads.Exclusions = z.GetExclusions()
                            STATIC_VARIABLES.CurrentExclusionSet = exclusionSet
                            'MsgBox(fname & " | " & lname & vbCrLf & "Email ? : " & canGetEmail & vbCrLf & "Email Address: " & emailAddy & vbCrLf & "Lead Num: " & leadNum & vbCrLf & vbCrLf & "Exclusions :>" & vbCrLf & "Generated: " & exclusionSet.Generated & vbCrLf & "Marketer: " & exclusionSet.Marketer & vbCrLf & "PLS: " & exclusionSet.PLS & " | SLS: " & exclusionSet.SLS & vbCrLf & "LastMResult:" & exclusionSet.LastMResult & vbCrLf & "Phone: " & exclusionSet.Phone, MsgBoxStyle.Information, "Debug Info:")
                            If canGetEmail = True Then
                                '' check here to make sure not called an cancelled
                                '' 8-26-2014

                                If z.IsCalledAndCancelled(leadNum, Me.dtpIssueLeads.Value.ToShortDateString) = False Then
                                    MsgBox("This lead was called and cancelled." & vbCrLf & "Please check your marketing results and try again.", MsgBoxStyle.Critical, "Lead Called and Cancelled")
                                    Exit Sub
                                End If



                                emlBody = z.ConstructMessageWithoutExclusions(fname, lname, leadNum, emailAddy)
                                'MsgBox(emlBody)
                                '' uncomment here to actually send
                                z.EMAIL_SINGLE_MarkupEmail_WITHOUT_EXCLUSIONS(fname, lname, leadNum, emailAddy, emlBody, "Record ID: " & leadNum.ToString)
                            ElseIf canGetEmail = False Then
                                '' do nothing for now. 
                            End If
                        End If

                        'End If



                    Next
                End If
            Next
        End If


    End Sub

    Private Sub tsAttachedFilesNAV_Click(sender As Object, e As EventArgs) Handles tsAttachedFilesNAV.Click

        'MsgBox(dir1 & vbCrLf & dir2)
        'Me.GetRidOfOldAndPutNew()

        Dim x As Control
        For Each x In Me.pnlAFPics.Controls
            If TypeOf (x) Is ListView Then
                If x.Name = "lsAF" Then
                    Dim a As ListView = x
                    Dim b As New ReusableListViewControl
                    b.ChangeDirectory(STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID, a)
                End If
            ElseIf x.Name = "lsJP" Then
                If x.Name = "lsJP" Then
                    Dim a As ListView = x
                    Dim b As New ReusableListViewControl
                    b.ChangeDirectory(STATIC_VARIABLES.JobPicturesFileDirectory & STATIC_VARIABLES.CurrentID, a)
                End If
            End If
        Next

    End Sub

    Private Sub GetRidOfOldAndPutNew()
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim height As Integer = 0
        Dim width As Integer = 0
        Dim name As String = ""
        Dim tg As String = ""

        Dim x2 As Integer = 0
        Dim y2 As Integer = 0
        Dim height2 As Integer = 0
        Dim width2 As Integer = 0
        Dim name2 As String = ""
        Dim tg2 As String = ""


        For Each Ctrl As Windows.Forms.Control In Me.pnlAFPics.Controls
            If TypeOf Ctrl Is ListView Then
                '' ' lsAF ' and ' lsJP ' 
                If Ctrl.Name = "lsAF" Then
                    x = Ctrl.Location.X
                    y = Ctrl.Location.Y
                    height = Ctrl.Height
                    width = Ctrl.Width
                    name = "lsAF"
                    tg = Ctrl.Tag
                End If
                If Ctrl.Name = "lsJP" Then
                    x2 = Ctrl.Location.X
                    y2 = Ctrl.Location.Y
                    height2 = Ctrl.Height
                    width2 = Ctrl.Width
                    name2 = "lsJP"
                    tg2 = Ctrl.Tag
                End If
            End If
        Next

        Dim ls As ListView = pnlAFPics.Controls("lsAF")
        pnlAFPics.Controls.Remove(ls)
        Dim dir1 As String = (STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID).ToString
        Dim dir2 As String = (STATIC_VARIABLES.JobPicturesFileDirectory & STATIC_VARIABLES.CurrentID).ToString
        Dim ls2 As ListView = pnlAFPics.Controls("lsJP")
        pnlAFPics.Controls.Remove(ls2)

        
        Dim widthOfParent As Integer = pnlAFPics.Width
        Dim widthOfControl As Integer = (widthOfParent / 2) - 20
        Dim heightOfParent As Integer = pnlAFPics.Height
        Dim heightOfControl As Integer = (heightOfParent - 30)
        Dim InitPoint As System.Drawing.Point = New System.Drawing.Point((0 + 10), (0 + 10))
        Dim InitPoint2 As System.Drawing.Point = New System.Drawing.Point(((widthOfParent / 2) + 10), (0 + 10))

        Dim rc As ReusableListViewControl = New ReusableListViewControl
        rc.GenerateListControl(pnlAFPics, dir1, InitPoint, "lsAF", heightOfControl, widthOfControl)
        Dim rc2 As ReusableListViewControl = New ReusableListViewControl
        rc2.GenerateListControl(pnlAFPics, dir2, InitPoint2, "lsJP", heightOfControl, widthOfControl)

    End Sub

    Private Sub btnPrintIssue_Click(sender As Object, e As EventArgs) Handles btnPrintIssue.Click
        'MsgBox("btnPrintIssue")
    End Sub


#Region "Printing Options"

    Private Sub btnPrintIssuedAppts_Click(sender As Object, e As EventArgs) Handles btnPrintIssuedAppts.Click
        MsgBox("btnPrintIssuedAppts")
    End Sub

    Private Sub btnPrintAllIssue_Click(sender As Object, e As EventArgs) Handles btnPrintAllIssue.Click
        '' MsgBox("btnPrintAllIssue")
        If Me.btnExclude.Text.Contains("On") Then
            frmPrint.Exclusions = False
            Dim y As Panel
            Dim arLeadNums As New ArrayList
            For Each y In pnlIssue.Controls
                Dim yy As Control
                For Each yy In y.Controls
                    If TypeOf yy Is LinkLabel Then
                        ''MsgBox("Record ID: " & yy.Text,information,"DEBUG INFO")
                        Dim lead_id As String = yy.Text
                        arLeadNums.Add(lead_id)
                    End If
                Next
            Next

            frmPrint.ClearListView()

            Dim i As Integer = 0
            For i = 0 To arLeadNums.Count - 1
                Dim lv As New ListViewItem
                lv.Text = arLeadNums(i)
                frmPrint.lsLeadIds.Items.Add(lv)
            Next
            frmPrint.ShowDialog()
        ElseIf Me.btnExclude.Text.Contains("Off") Then
            frmPrint.Exclusions = True
            Dim y As Panel
            Dim arLeadNums As New ArrayList
            For Each y In pnlIssue.Controls
                Dim yy As Control
                For Each yy In y.Controls
                    If TypeOf yy Is LinkLabel Then
                        ''MsgBox("Record ID: " & yy.Text,information,"DEBUG INFO")
                        Dim lead_id As String = yy.Text
                        arLeadNums.Add(lead_id)
                    End If
                Next
            Next

            frmPrint.ClearListView()

            Dim i As Integer = 0
            For i = 0 To arLeadNums.Count - 1
                Dim lv As New ListViewItem
                lv.Text = arLeadNums(i)
                frmPrint.lsLeadIds.Items.Add(lv)
            Next
            frmPrint.ShowDialog()
        End If
        

    End Sub

    Private Sub btnPrintApptSheet_Click(sender As Object, e As EventArgs) Handles btnPrintApptSheet.Click
        MsgBox("btnPrintApptSheet")
    End Sub

    Private Sub btnPrintCurrentList_Click(sender As Object, e As EventArgs) Handles btnPrintCurrentList.Click
        MsgBox("btnPrintCurrentList")
    End Sub

    Private Sub btnPrintNoEmailIssue_Click(sender As Object, e As EventArgs) Handles btnPrintNoEmailIssue.Click
        ''MsgBox("btnPrintNoEmailIssue")
        If Me.btnExclude.Text.Contains("On") Then
            frmPrint.Exclusions = False
            Dim y As Panel
            Dim arRepNames As New ArrayList
            Dim arLeadNums As New ArrayList
            Dim _leadNum As String = ""
            For Each y In pnlIssue.Controls
                Dim ctrl As Control
                For Each ctrl In y.Controls

                    If TypeOf ctrl Is LinkLabel Then
                        _leadNum = ctrl.Text
                    End If
                    If TypeOf ctrl Is ComboBox Then
                        If ctrl.Text <> "" Then
                            Dim strName = Split(ctrl.Text, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                            Dim fname As String = strName(0)
                            Dim lname As String = strName(1)
                            Dim g As New bulkPrintOperations
                            Dim canGetMail As Boolean = g.CanRepGetEmail(fname, lname)
                            If canGetMail = False Then
                                arRepNames.Add(fname & " " & lname)
                                arLeadNums.Add(_leadNum)
                            End If
                        End If
                    End If
                Next
            Next

            Dim gg As Integer = 0
            ''Dim strMSG As String = ""
            frmPrint.ClearListView()

            For gg = 0 To arRepNames.Count - 1
                '' strMSG += arLeadNums(gg) & " " & arRepNames(gg) & vbCrLf
                Dim lv As New ListViewItem
                lv.Text = arLeadNums(gg)
                frmPrint.lsLeadIds.Items.Add(lv)
            Next

            frmPrint.ShowDialog()


        ElseIf btnExclude.Text.Contains("Off") Then
            '' MsgBox("exclusions on")
            frmPrint.Exclusions = True
            Dim b As New bulkPrintOperations
            Dim ex_set As bulkPrintOperations.Exclusions
            ex_set = b.GetExclusions()
            Dim y As Panel
            Dim arRepNames As New ArrayList
            Dim arLeadNums As New ArrayList
            Dim _leadNum As String = ""
            For Each y In pnlIssue.Controls
                Dim ctrl As Control
                For Each ctrl In y.Controls

                    If TypeOf ctrl Is LinkLabel Then
                        _leadNum = ctrl.Text
                    End If
                    If TypeOf ctrl Is ComboBox Then
                        If ctrl.Text <> "" Then
                            Dim strName = Split(ctrl.Text, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                            Dim fname As String = strName(0)
                            Dim lname As String = strName(1)
                            Dim g As New bulkPrintOperations
                            Dim canGetMail As Boolean = g.CanRepGetEmail(fname, lname)
                            If canGetMail = False Then
                                arRepNames.Add(fname & " " & lname)
                                arLeadNums.Add(_leadNum)
                            End If
                        End If
                    End If
                Next
            Next
            Dim gg As Integer = 0
            ''Dim strMSG As String = ""
            frmPrint.ClearListView()

            For gg = 0 To arRepNames.Count - 1
                '' strMSG += arLeadNums(gg) & " " & arRepNames(gg) & vbCrLf
                Dim lv As New ListViewItem
                lv.Text = arLeadNums(gg)
                frmPrint.lsLeadIds.Items.Add(lv)
            Next


            frmPrint.ShowDialog()
        End If


    End Sub

    Private Sub btnPrintThisIssue_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintThisIssue.Click
        ''
        '' print selected lead 
        '' 
        ''MsgBox("btnPrintThisIssue")
        ''
        If btnExclude.Text.Contains("On") Then
            frmPrint.Exclusions = False
            Dim y As Panel
            For Each y In pnlIssue.Controls
                If y.BorderStyle = BorderStyle.FixedSingle Then
                    Dim yy As Control
                    For Each yy In y.Controls
                        If TypeOf yy Is LinkLabel Then
                            ''MsgBox("Record ID: " & yy.Text,information,"DEBUG INFO")
                            Dim lead_id As String = yy.Text
                            Dim lv As New ListViewItem
                            lv.Text = lead_id
                            frmPrint.ClearListView()
                            frmPrint.lsLeadIds.Items.Add(lv)
                            frmPrint.ShowDialog()

                        End If
                    Next
                End If
            Next
        ElseIf btnExclude.Text.Contains("Off") Then
            frmPrint.Exclusions = True
            Dim y As Panel
            For Each y In pnlIssue.Controls
                If y.BorderStyle = BorderStyle.FixedSingle Then
                    Dim yy As Control
                    For Each yy In y.Controls
                        If TypeOf yy Is LinkLabel Then
                            ''MsgBox("Record ID: " & yy.Text,information,"DEBUG INFO")
                            Dim lead_id As String = yy.Text
                            Dim lv As New ListViewItem
                            lv.Text = lead_id
                            frmPrint.ClearListView()
                            frmPrint.lsLeadIds.Items.Add(lv)
                            frmPrint.ShowDialog()
                        End If
                    Next
                End If
            Next
        End If
        
    End Sub

#End Region



#Region "Email Templating Stuff"
    Public Function GetTemplates()

        Dim y As New emlTemplateLogic
        Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim depart_ As String = y.GetEmployeeDepartment(name(0), name(1), False)
        Dim g As List(Of emlTemplateLogic.TemplateInfo)
        g = y.GetTemplatesByDepartment(False, depart_)
        Dim a As emlTemplateLogic.TemplateInfo
        Return g

    End Function
#End Region
     
    Private Sub btnEMailCustomer_Click(sender As Object, e As EventArgs) Handles btnEMailCustomer.Click
        
        ' MsgBox("Hit Test", MsgBoxStyle.Information, "DEBUG INFO")
        '' pull up frmEmailPreview
        '' 3 public vars
        'Public LeadToShow As convertLeadToStruct.EnterLead_Record
        'Public TemplateName As String = ""
        'Public Department As String = ""
        Dim obj_ As New convertLeadToStruct
        Dim y As New emlTemplateLogic
        Dim _lead As New convertLeadToStruct.EnterLead_Record
        Dim recID As String = STATIC_VARIABLES.CurrentID.ToString
        If Len(recID) <= 0 Then
            '' take first lead in dbase | testing 
            recID = y.GetMaxID(False)
        ElseIf Len(recID) >= 1 Then
            '' use current recID
            recID = recID
        End If
        _lead = obj_.ConvertToStructure(recID, False)
        frmEmailPreview.LeadToShow = _lead
        Dim name() = Split(STATIC_VARIABLES.CurrentUser, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim depart_ As String = y.GetEmployeeDepartment(name(0), name(1), False)

        frmEmailTemplateChoice.ShowDialog()





    End Sub

    Private Sub dtpSummary2_LostFocus(sender As Object, e As EventArgs) Handles dtpSummary2.LostFocus
        focusdtp2 = False
    End Sub

    Private Sub dtpSummary2_ValueChanged(sender As Object, e As EventArgs) Handles dtpSummary2.ValueChanged
        If LoadComplete = False Then
            Exit Sub
        End If
        If focusdtp2 = True Then
            If dtpsum2orig <> Me.dtpSummary2.Value.ToString Then
                Me.cboDateRangeSummary.SelectedItem = "Custom"
                dtpsum2orig = Me.dtpSummary2.Value.ToString
                cboDateRangeSummary_SelectedIndexChanged(Nothing, Nothing)
            End If
        End If

    End Sub

    Private Sub dtpSummary_LostFocus(sender As Object, e As EventArgs) Handles dtpSummary.LostFocus
        Focusdtp1 = False
    End Sub

    Private Sub dtpSummary_ValueChanged(sender As Object, e As EventArgs) Handles dtpSummary.ValueChanged
        If LoadComplete = False Then
            Exit Sub
        End If
        If Focusdtp1 = True Then
            If dtpsum1orig <> Me.dtpSummary.Value.ToString Then
                Me.cboDateRangeSummary.SelectedItem = "Custom"
                Me.dtpsum1orig = Me.dtpSummary.Value.ToString
                cboDateRangeSummary_SelectedIndexChanged(Nothing, Nothing)
            End If
        End If
   
    End Sub

    Private Sub tbMain_SizeChanged(sender As Object, e As EventArgs) Handles tbMain.SizeChanged
        'Dim x As Integer = (Me.tbMain.Height / 5)
        'MsgBox(x.ToString)
        'Me.tbMain.ItemSize = New Size(x, 20)
    End Sub
End Class
