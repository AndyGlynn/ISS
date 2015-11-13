Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class ALERT_LOGIC
   

#Region "Notes"
    '' Notes:
    ''
    '' Need a method to gather alerts and list successively
    '' Either dump to a dataset, or bypass the data set and create a class that dynamically creates a panel,
    '' with nested controls to nest in the pnlBack control on the static form.
    '' Check for alerts on load of program, IE look for alerts that are past due and have not been dealt with
    '' likewise, show those alerts in a panel view
    '' that simulates some kind of listview with icons
    '' need a method to push back the alert for five minutes
    '' need a method to toggle weather or not the alert has been dealt with / completed.
    ''
    '' GENERAL FLOW:
    '' 1) count alerts that are not completed and = to time it is
    '' 2) is count greater than 1?
    '' 2a) yes. Get an array/dataset to house all the records.
    '' 2b) no. exit
    '' 3) for each record in array/dataset, create a custom panel preformatted and add to pnlback
    '' 4) on pnlback house a variable for the selected item
    '' 5) either, toggle completed, or push back five minutes
    '' 6) end of flow.
    '' 7) close form. (alerts dealt with or not - alerts not dealt with recycled on load of program until dealt with)
    '' 
    '' Time Frames:
    '' in order to pick up the alert, need a 15 second window +/- to make sure it is caught due to the timer firring off
    '' on odd clock cycles.
    '' in order to push back an alert, utilize a timespan() method and add five (5) minutes to it.
    '' 
    '' Panels:
    '' need a mouseover effect to show the panel is highlighted
    '' need to add a hyperlink label to quick bounce to the record selected
    '' need to pull record ID out and hide invisibly for ease of manipulation of said recordset IE panel.tag = UniqueRecordID
    '' panel needs an on_click event to get the Record ID out to form variable of recordID
    '' form needs an image list to show what type of alert it is
    '' need a refresh method for real time effect of showing alert has been dealt with in panels
    '' alternating panels should have an off color to simulate a financial ledger for ease of readability
    '' 


    '' Table Structure: iss.dbo.AlertTable
    '' |ID|LeadNum|UserName|AlertTime|ExecutionDate|Notes|AssignedBy|Completed|
    ''  0   1       2        3         4             5      6          7
    '' 
    '' Stored Proc: INSERT
    ''
    ''set ANSI_NULLS ON
    ''set QUOTED_IDENTIFIER ON
    ''GO
    ''ALTER procedure [dbo].[InsertAlert](
    ''@LeadNum nvarchar(50),
    ''@UserName nvarchar(150),
    ''@AlertTime datetime,
    ''@ExecutionDate datetime,
    ''@Notes nvarchar(max),
    ''@AssignedBy nvarchar(150),
    ''@Completed bit
    '')
    ''as
    ''set nocount off;
    ''INSERT iss.dbo.AlertTable(LeadNum,UserName,AlertTime,ExecutionDate,Notes,AssignedBy,Completed)
    ''values(@LeadNum,@UserName,@AlertTime,@ExecutionDate,@Notes,@AssignedBy,@Completed)

    '' Stored Proc: UPDATE
    ''
    ''create proc dbo.UpdateAlert
    ''(
    ''@RecID nvarchar(max),
    ''@ID nvarchar(max),
    ''@UserName nvarchar(max),
    ''@AlertTime datetime,
    ''@ExecutionDate datetime,
    ''@Notes nvarchar(max),
    ''@AssignedBy nvarchar(max),
    ''@Completed bit
    '')
    ''as set nocount off;
    ''Update iss.dbo.AlertTable
    ''Set LeadNum = @ID,
    ''	UserName = @UserName,
    ''	AlertTime = @AlertTime,
    ''	ExecutionDate = @ExecutionDate,
    ''	Notes = @Notes,
    ''	AssignedBy = @AssignedBy,
    ''	Completed = @completed
    ''WHERE ID = @RecID
    ''GO

    '' Stored Proc: DELETE
    ''
    ''create proc dbo.DeleteAlert
    ''(
    ''@RecID nvarchar(max)
    '')
    ''as
    ''set nocount off;
    ''Delete iss.dbo.alerttable
    ''where ID = @RecID
    ''GO

    '' Stored Proc: COUNT / Not Completed / need to be dealt with
    ''create proc dbo.CountAlert
    ''as
    ''set nocount off;
    ''SELECT COUNT(ID) from iss.dbo.AlertTable where Completed = 0
    ''GO

#End Region
#Region "Variables"
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Private cntOfAlerts As Integer = 0

    Public TargetID As String = ""
    '' |ID|LeadNum|UserName|AlertTime|ExecutionDate|Notes|AssignedBy|Completed|
    ''  0   1       2        3         4             5      6          7
    Public dset_Alerts As Data.DataSet = New Data.DataSet("DSET_ALERTS")
    Public dtab_Alerts As Data.DataTable = New Data.DataTable("tblAlerts")
    Public dcol_ID As Data.DataColumn = New Data.DataColumn("ID")
    Public dcol_LeadNum As Data.DataColumn = New Data.DataColumn("LeadNUM")
    Public dcol_UserName As Data.DataColumn = New Data.DataColumn("UserName")
    Public dcol_AlertTime As Data.DataColumn = New Data.DataColumn("AlertTime")
    Public dcol_ExecutionDate As Data.DataColumn = New Data.DataColumn("ExecutionDate")
    Public dcol_Notes As Data.DataColumn = New Data.DataColumn("Notes")
    Public dcol_AssignedBy As Data.DataColumn = New Data.DataColumn("AssignedBy")
    Public dcol_Completed As Data.DataColumn = New Data.DataColumn("Completed")
    Public dcol_CustFName As Data.DataColumn = New Data.DataColumn("CustFName")
    Public dcol_CustLName As Data.DataColumn = New Data.DataColumn("CustLName")
    Public dcol_Phone As Data.DataColumn = New Data.DataColumn("Phone")

#End Region
#Region "Methods"
    



    Public Sub GetRecordInformation(ByVal UserName As String)
        '' for right now, just get all the records to plot them out
        '' will need to edit this call to represent time frames
        Try
            Dim cmdGET As SqlCommand = New SqlCommand _
            ("SELECT AlertTable .*, EnterLead .Contact1FirstName ,EnterLead .Contact1LastName ,EnterLead .HousePhone  from iss.dbo.alerttable inner join EnterLead  on enterlead.id = alerttable.LeadNUM   where completed = 0 and UserName = @USR and (select ExecutionDate + AlertTime)  <= {fn current_timestamp()} order by ExecutionDate asc , AlertTime asc", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@USR", UserName)
            cmdGET.Parameters.Add(param1)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            While r1.Read
                Dim r As Data.DataRow = dtab_Alerts.NewRow
                r("ID") = r1.Item(0)
                r("LeadNUM") = r1.Item(1)
                r("UserName") = r1.Item(2)
                Dim time = Split(r1.Item(3).ToString, " ", 2)
                r("AlertTime") = time(1).ToString
                Dim d1 = Split(r1.Item(4).ToString, " ", 2)
                r("ExecutionDate") = d1(0).ToString
                r("Notes") = r1.Item(5)
                r("AssignedBy") = r1.Item(6)
                r("Completed") = r1.Item(7)
                r("CustFName") = r1.Item(8)
                r("CustLName") = r1.Item(9)
                r("Phone") = r1.Item(10)
                dtab_Alerts.Rows.Add(r)
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ALERT_LOGIC", "UserName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm.ToString, "SQL", "GetRecordInformation")

        End Try

    End Sub
#End Region
#Region "Properties"
    Public Property CountOfAlerts() As Integer
        Get
            Return cntOfAlerts
        End Get
        Set(ByVal value As Integer)
            cntOfAlerts = value
        End Set
    End Property


#End Region
#Region "Functions"
    Public Function CountAlerts(ByVal UserName As String)
        Try
            Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.AlertTable where Completed = 0 and UserName = @USR and (select ExecutionDate + AlertTime)  <= {fn current_timestamp()}", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@USR", UserName)
            cmdCNT.Parameters.Add(param1)
            cmdCNT.CommandType = CommandType.Text
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdCNT.ExecuteReader
            While r1.Read
                CountOfAlerts = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
            Return CountOfAlerts
        Catch ex As Exception
            cnn.Close()
            Return CountOfAlerts
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ALERT_LOGIC", "UserName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "CountAlerts")
        End Try
    End Function
  
 
#End Region
#Region "New Object"
    Public Sub New(ByVal UserName As String)
        Try
            dset_Alerts.Tables.Add(dtab_Alerts)
            dtab_Alerts.Columns.Add(dcol_ID) '0
            dtab_Alerts.Columns.Add(dcol_LeadNum) '1
            dtab_Alerts.Columns.Add(dcol_UserName) '2
            dtab_Alerts.Columns.Add(dcol_AlertTime) '3
            dtab_Alerts.Columns.Add(dcol_ExecutionDate) '4
            dtab_Alerts.Columns.Add(dcol_Notes) '5
            dtab_Alerts.Columns.Add(dcol_AssignedBy) '6
            dtab_Alerts.Columns.Add(dcol_Completed) '7
            dtab_Alerts.Columns.Add(dcol_CustFName) '8
            dtab_Alerts.Columns.Add(dcol_CustLName) '9 
            dtab_Alerts.Columns.Add(dcol_Phone)     '10 

            Dim cnt As Integer = CountAlerts(UserName)
            Me.CountOfAlerts = cnt
            Select Case cnt
                Case Is >= 1
                    GetRecordInformation(UserName)
                    ' due to desing changes, this is going to be switched over to a 
                    ' simulated paging mechanism on 'ManageAlerts'
                    ' the revised flow is:
                    ' 1) count alerts. 
                    ' 2) is it >=1?
                    ' 2a) yes. show form and pass it the dataset object.
                    ' 2b) no. exit / continue to log in.
                    ' 3) either toggle alert complete or set new date time
                    ' 4) scroll to next record in dataset
                    ' 5) end.
                    'ManageAlerts.MdiParent = Main
                    'ManageAlerts.Visible = False
                    'ManageAlerts.Show()

                    ManageAlerts.dset_Alerts = Me.dset_Alerts

                    Exit Select
                Case Is <= 0
                    Exit Select
            End Select
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ALERT_LOGIC", "UserName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "'New'")
            cnn.Close()
        End Try

    End Sub
#End Region
#Region "Nested Classes"
    Public Class CreateForePanel
        '' the class to create the panel with respective information
        '' 
        Friend WithEvents pctImage As New PictureBox
        'Friend WithEvents lblUserName As Label  || not needed, assumed the user logged in is the user the alerts are for.
        Friend WithEvents lblAlertTime As New Label
        Friend WithEvents lblExecDate As New Label
        Friend WithEvents lblNotes As New Label
        Friend WithEvents lblAssignedBy As New Label
        Friend WithEvents lnkID As New LinkLabel
        Friend WithEvents ForePanel As New Panel
        Public F1 As New Font("Tahoma", 8.25!, FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte))
        Public Sub New(ByVal ID As String, ByVal LeadNumber As String, ByVal AlertTime As String, ByVal ExecDate As String, ByVal Notes As String, ByVal AssignedBy As String, ByVal Completed As Boolean)
            Try
                '' panel to nest below controls
                'ForePanel = New Panel
                ForePanel.BackColor = Color.White
                ForePanel.Tag = ID
                ForePanel.Size = New Size(626, 29)
                ForePanel.Location = New Point(3, 3)

                '' picturebox for visual styles
                'pctImage = New PictureBox
                pctImage.Size = New Size(18, 18)
                pctImage.Location = New Point(12, 4)
                pctImage.SizeMode = PictureBoxSizeMode.CenterImage
                pctImage.Image = PastDueAlerts.ImageList1.Images.Item(0)

                '' User name label || Not needed
                'lblUserName = New Label
                'lblUserName.Font = F1
                'lblUserName.Text = UserName.ToString
                'lblUserName.Size = New Size()

                '' Alert Time
                'lblAlertTime = New Label
                lblAlertTime.AutoEllipsis = True
                lblAlertTime.Font = F1
                lblAlertTime.Text = AlertTime.ToString
                lblAlertTime.Size = New Size(90, 13)
                lblAlertTime.Location = New Point(86, 8)

                '' Execution Date
                'lblExecDate = New Label
                lblExecDate.AutoEllipsis = True
                lblExecDate.Font = F1
                lblExecDate.Text = ExecDate.ToString
                lblExecDate.Size = New Size(97, 13)
                lblExecDate.Location = New Point(197, 8)

                '' Notes Label
                'lblNotes = New Label
                lblNotes.AutoEllipsis = True
                lblNotes.Font = F1
                lblNotes.Text = Notes
                lblNotes.Size = New Size(189, 13)
                lblNotes.Location = New Point(330, 8)

                '' Assigned By label
                'lblAssignedBy = New Label
                lblAssignedBy.AutoEllipsis = True
                lblAssignedBy.Text = AssignedBy
                lblAssignedBy.Size = New Size(79, 13)
                lblAssignedBy.Location = New Point(550, 8)

                '' Link ID 
                'lnkID = New LinkLabel
                lnkID.Font = F1
                lnkID.Text = LeadNumber
                lnkID.Location = New Point(34, 8)
                lnkID.Size = New Size(37, 13)

                ForePanel.Controls.Add(pctImage)
                ForePanel.Controls.Add(lnkID)
                ForePanel.Controls.Add(lblAlertTime)
                ForePanel.Controls.Add(lblExecDate)
                ForePanel.Controls.Add(lblNotes)
                ForePanel.Controls.Add(lblAssignedBy)
            Catch ex As Exception
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ALERT_LOGIC.CreateForePanel", "ByVal ID As String, ByVal LeadNumber As String, ByVal AlertTime As String, ByVal ExecDate As String, ByVal Notes As String, ByVal AssignedBy As String, ByVal Completed As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "'New'")
            End Try

        End Sub


        Private Sub ForePanel_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ForePanel.MouseEnter
            'Me.ForePanel.BackColor = Color.Blue
        End Sub

        Private Sub ForePanel_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ForePanel.MouseLeave
            'Me.ForePanel.BackColor = Color.White
        End Sub
        Private Sub forepanel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ForePanel.Click

            If Me.ForePanel.BorderStyle = BorderStyle.FixedSingle Then
                Me.ForePanel.BorderStyle = BorderStyle.None
                Me.ForePanel.BackColor = Color.White
                Dim indx As Integer = PastDueAlerts.SelectedID.IndexOf(Me.ForePanel.Tag)
                PastDueAlerts.SelectedID.RemoveAt(indx)
                'Form1.SelectedID.RemoveAt(Form1.SelectedID.Count - 1)
            ElseIf Me.ForePanel.BorderStyle = BorderStyle.None Then
                Me.ForePanel.BorderStyle = BorderStyle.FixedSingle
                Me.ForePanel.BackColor = Color.SlateGray
                PastDueAlerts.SelectedID.Add(Me.ForePanel.Tag.ToString)
            End If

            'MsgBox("Unique Record ID# " & Form1.SelectedID.ToString)
        End Sub

        Private Sub lnkID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkID.Click
            'MsgBox("Show information for ID#: " & Me.lnkID.Text.ToString)
            CommonNFO.RecordID = Me.lnkID.Text
            CommonNFO.ShowInTaskbar = False
            CommonNFO.ShowDialog()
        End Sub
    End Class
#End Region
   
End Class
