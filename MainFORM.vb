

Imports System
Imports System.Threading
Imports System.Text
Imports System.Net
Imports System.Windows.Forms
Imports System.IO
Imports System.Net.Sockets
Imports System.Text.StringBuilder
Imports System.Windows
Imports Microsoft.VisualBasic
Imports Microsoft
Imports System.Windows.Forms.Form
Imports Microsoft.VisualBasic.Interaction
Imports System.Drawing

Public Class Main
    Friend WithEvents CloseAll As New ToolStripMenuItem
    Friend WithEvents TileVertically As New ToolStripMenuItem
    Friend WithEvents TileHorizontally As New ToolStripMenuItem
    Friend WithEvents Cascade As New ToolStripMenuItem

    Public NetListener As TcpListener '' listen on port # 5586 | sends on 5587
    Public NetListener2 As TcpListener '' listen on port # 5587 | sends Accept/Decline info
    Public ThNet As New Thread(AddressOf DoListen)
    Public ThAD As New Thread(AddressOf DoListen2)
    Public Msg As String = ""
    Public Event GotMessage(ByVal Msg As String) '' check incomming calls
    Public Event GotMessage2(ByVal Msg As String) '' event for accept/decline
    Friend WithEvents btnClose As New Button ' button for msg form to close
    Public Delegate Sub DisplayInvoker(ByVal t As String)
    Public LeadNumC As Integer = 0
    Public LineC As String = ""
    Public WhoC As String = ""
    Public AutoC As String = ""
    Public NoteC As String = ""
    Public forcettconfirm As Boolean = False
    Public forcettsales As Boolean = False
    Private mobjClient As TcpClient
    Private marData(1024) As Byte
    Private mobjText As New StringBuilder()
    Friend WithEvents frmbutton As ToolStripMenuItem
    Public Function GetIPv4Address()
        GetIPv4Address = String.Empty
        Dim strmachine As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strmachine)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetworkV6 Then
                GetIPv4Address = ipheal
                ' MsgBox(ipheal.ToString, MsgBoxStyle.Critical, "ERROR")

            End If
        Next

    End Function
    Private Sub tsbsales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbsales.Click
        Dim x As Form
        x = Sales
        'sales.LOAD_HANDLER = "LOAD"
        'Sales.lvAttachedFiles.Size = New System.Drawing.Size(Sales.pnlAFPics.Width / 2 - 5, Sales.lvAttachedFiles.Height)
        'Sales.lvJobPics.Size = New System.Drawing.Size(Sales.pnlAFPics.Width / 2 - 5, Sales.lvJobPics.Height)
        Sales.Label4.Location = New System.Drawing.Point((Sales.tpSummary.Width / 2) - 87, Sales.Label4.Location.Y)
        x.MdiParent = Me
        x.Show()
        x.BringToFront()

    End Sub

    Private Sub tsbnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbnew.Click
        '' put this back to show enterlead inside of mdiparent
        '' 'Main'

        'Dim y As New Form
        'y = My.Forms.EnterLead

        EnterLead.MdiParent = Me
        If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
            STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
        End If


        EnterLead.Show()
        EnterLead.StartPosition = FormStartPosition.CenterParent
        EnterLead.WindowState = FormWindowState.Normal
        'EnterLead.PerformLayout()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        Application.Exit()
    End Sub

    Private Sub tsbattach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbattach.Click
        AttachAFile.ShowInTaskbar = False
        AttachAFile.ShowDialog()
        'Try
        '    Dim y As String = InputBox$("Please enter the lead number you wish to attach the file(s) to.", "Enter Lead Number", STATIC_VARIABLES.CurrentID.ToString) ' RecordLogic.CurrentID was taken out for default recordID in rev 5
        '    If y = "" Then
        '        MsgBox("You must supply a lead number to attach a file to.", MsgBoxStyle.Exclamation, "Error Attaching File")
        '        Exit Sub
        '    End If
        '    Dim z As New Attach
        '    Dim a As String = "0"
        '    z.AttachFile(y, a)
        '    'MsgBox("File attached to Lead#: " & y.ToString & "", MsgBoxStyle.Information, "File Attached Successfully")
        '    Exit Sub
        'Catch ex As Exception
        '    MsgBox(ex.Message.ToString)
        'End Try
    End Sub

    Private Sub tsImportsPics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsImportsPics.Click
        Dim x As Form
        x = ImportPictures
        'x.MdiParent = Me
        x.ShowInTaskbar = False
        x.ShowDialog()
    End Sub

    Private Sub tsbrolodex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbrolodex.Click
        Dim x As Form
        x = frmRolodex 'Employee_Contacts
        'x.MdiParent = Me
        x.ShowInTaskbar = False
        x.ShowDialog()
    End Sub

    Private Sub tsbtransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtransfer.Click
        'Dim x As Form
        'x = Transfer_Lead
        'x.ShowInTaskbar = False
        'x.ShowDialog()
    End Sub

    Private Sub Main_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()

    End Sub

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'NetListener.Stop() '' caller listener shutdown
        'NetListener2.Stop() '' accept/decline listener shutdown
        'Dim c As New USER_LOGIC
        'c.LogOut_Of_System(STATIC_VARIABLES.Login, STATIC_VARIABLES.UserPWD)
        Dim x As New MAPPOINT_LOGIC
        x.KillProcess() '' make sure all mappoint instances are not runnning and closed out / free resources 
        Dim c As New USER_LOGICv2
        Try
            Dim frm As Windows.Forms.Form = Me.ActiveMdiChild
            Dim scrX As Integer = frm.Location.X
            Dim scrY As Integer = frm.Location.Y
            Dim scrW As Integer = frm.Width
            Dim scrH As Integer = frm.Height
            c.LogOut("Aaron", "Clay", scrX, scrY, scrW, scrH, frm.Name.ToString, "No Query", My.Computer.Name.ToString, "admin")
        Catch ex As Exception
            Dim scrX As Integer = 0
            Dim scrY As Integer = 0
            Dim scrW As Integer = 0
            Dim scrH As Integer = 0
            c.LogOut("99", scrX, scrY, scrW, scrH, "Main", "No Query", My.Computer.Name.ToString, "admin")
        End Try


    End Sub

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.tmrAlerts.Stop()

        'ManageAlerts.MdiParent = Me


        'ThNet.Start()
        'ThAD.Start()
        Me.Text = STATIC_VARIABLES.ProgramName
        Me.TSMain.Enabled = True

        tmrAlerts.Enabled = True
        'tmrXFER.Enabled = True
        'tmrStartupLauncher.Enabled = True







        'Dim r As New REG_LOGIC
        'r.ReadKey(STATIC_VARIABLES.LicenseKey, STATIC_VARIABLES.LeaseKey)
        LOGIN.TopMost = True
        LOGIN.ShowDialog()
        If LOGIN.Success = False Then
            Exit Sub
        End If
        Dim g As New systemCheckLogin
        Dim cntRecs As Integer = g.CntRecs
        Me.tsLoggedInAs.Text = STATIC_VARIABLES.CurrentUser
        Me.tsRecCount.Text = cntRecs.ToString
        Me.tsProgress.Minimum = 1
        Me.tsProgress.Maximum = cntRecs
        Me.tsProgress.Value = cntRecs

        If g.CheckNetworkStatus = False Then
            Me.tsNetworkStatus.ForeColor = Color.Red
            Me.tsNetworkStatus.Text = "DOWN"
        ElseIf g.CheckNetworkStatus = True Then
            Me.tsNetworkStatus.ForeColor = Color.Green
            Me.tsNetworkStatus.Text = "UP"
        End If

        If g.CheckSQLStatus = True Then
            Me.tsSqlSTatus.ForeColor = Color.Green
            Me.tsSqlSTatus.Text = "UP"
        ElseIf g.CheckSQLStatus = False Then
            Me.tsSqlSTatus.ForeColor = Color.Red
            Me.tsSqlSTatus.Text = "DOWN"
        End If
        tmrAlerts.Start()
        ' MsgBox(STATIC_VARIABLES.CurrentUser)
        Dim b As New ALERT_LOGIC(STATIC_VARIABLES.CurrentUser)
        Dim hour As Integer = Now.Hour
        Dim minute As Integer = Now.Minute
        Dim seconds As Integer = Now.Second
        Dim correctedTime As String = Date.Today.Date & " " & hour & ":" & minute & ":" & seconds
        Select Case b.CountOfAlerts
            Case Is >= 1
                ManageAlerts.ShowInTaskbar = False
                'ManageAlerts.Visible = False
                ManageAlerts.Show()
                ManageAlerts.BringToFront()


                Exit Select
            Case Is <= 0
                Exit Select
        End Select


        If LOGIN.txtUserName.Text <> "" Then

            LOGIN.txtPWD.Focus()
        Else
            LOGIN.txtUserName.Focus()
        End If

    End Sub
    Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        Dim r As New REG_LOGIC

        'LOGIN.ShowDialog()
        'r.ReadKey(STATIC_VARIABLES.LicenseKey, STATIC_VARIABLES.LeaseKey)
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        'Dim c As New USER_LOGIC
        'c.LogOut_Of_System(STATIC_VARIABLES.CurrentUser, STATIC_VARIABLES.UserPWD)

    End Sub

    Private Sub tsbschedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbschedule.Click
        Dim x As Form
        x = ScheduleAction
        x.ShowInTaskbar = False
        x.ShowDialog()

        x.Focus()
        x.TopMost = True

    End Sub

    Private Sub tsbalert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbalert.Click
        Dim x As Form
        x = AssignAlert
        'x.MdiParent = Me
        x.ShowInTaskbar = False
        x.ShowDialog()
    End Sub

    Private Sub tmrAlerts_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrAlerts.Tick
        Dim User As String = STATIC_VARIABLES.CurrentUser

        Dim x As New Alert_Logic_Tick
        x.CountAlertsTick(User)

        Select Case x.CountOfAlertsTick
            Case Is = 1

                If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
                    STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
                End If
                x.Get_ID(User)
            Case Is >= 2
                If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
                    STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
                End If
                x.Get_ID(User)
                Me.tmrAlerts_Tick(Nothing, Nothing)





            Case Is <= 0

                Exit Select
        End Select


    End Sub
    Private Sub DoListen()
        'Try
        '    AddHandler GotMessage, AddressOf DisplayNetMsg
        '    Control.CheckForIllegalCrossThreadCalls = False
        '    Dim port As Int32 = 5586 '' currently unassigned by IANA as of 2-17-2007
        '    Dim ip As IPAddress '= Me.GetIPv4Address.ToString                                        ' holder for local machine name.
        '    Dim strMachineName As String = ""
        '    strMachineName = My.Computer.Name
        '    Dim ipHost As IPHostEntry
        '    ipHost = Dns.GetHostEntry(strMachineName)
        '    Dim ipAddr() As IPAddress = ipHost.AddressList
        '    ip = System.Net.IPAddress.Any ''ip = GetIPv4Address()
        '    Dim ep As IPEndPoint = New IPEndPoint(ip, port)
        '    NetListener = New TcpListener(ep)
        '    Try
        '        'Do
        '        NetListener.Start()
        '        Dim clib As New TcpClient
        '        clib = NetListener.AcceptTcpClient
        '        Dim NetStream As NetworkStream = clib.GetStream
        '        Dim bytes(clib.ReceiveBufferSize) As Byte
        '        NetStream.Read(bytes, 0, CInt(clib.ReceiveBufferSize))
        '        Dim cliData As String = Encoding.ASCII.GetString(bytes)
        '        Msg = cliData.ToString
        '        Dim machine As String = "AndyGlynn2"
        '        Dim Ip2 As String = "192.168.1.8"
        '        Dim msg2 As String = "Go Fuck Yourself"
        '        Dim CMD As String = "LOAD"
        '        Dim USR As String = "AndyG"
        '        Dim Rec As String = "Server"











        'Dim lm(Nothing) '' split off leadnum
        'Dim who '' split off requester
        'Dim line '' split off line holding on
        'Dim response '' split off auto response
        'Dim nts '' split off notes     ANDYGLYNN2|192.168.1.8| |LOAD|AndyG| |
        ''lm = Split(Msg, "|", 2)
        'LeadNumC = CType(lm(0), Integer) '' lead number
        'who = Split(lm(1), "|", 2)
        'WhoC = "AndyGlynn2"                                 'CType(who(0).ToString, String) '' person requesting"
        'line = Split(who(1), "|", 2)
        'LineC = CType(line(0), String)
        'response = Split(line(1), "|", 2)
        'AutoC = CType(response(0), String)
        'nts = Split(response(1), "|", 2)
        'NoteC = CType(nts(0), String)
        'RaiseEvent GotMessage(LeadNumC)
        'clib.Close()
        'NetListener.Stop()
        '' Loop Until False
        '    Catch ex As Exception
        '    MsgBox(ex.Message, , )

        'End Try
        'Catch ex As Exception
        '    MsgBox(ex.Message, , )
        'End Try
    End Sub
    Private Sub DoListen2()
        'Try
        '    AddHandler GotMessage2, AddressOf DisplayAcceptDecline
        '    Control.CheckForIllegalCrossThreadCalls = False
        '    Dim port As Int32 = 5587 '' currently unassigned by IANA as of 2-17-2007
        '    Dim ip As IPAddress ' holder for local machine name.
        '    Dim strMachineName As String = ""
        '    strMachineName = My.Computer.Name
        '    Dim ipHost As IPHostEntry
        '    ipHost = Dns.GetHostByName(strMachineName)
        '    Dim ipAddr() As IPAddress = ipHost.AddressList
        '    ip = GetIPv4Address()
        '    NetListener2 = New TcpListener(ip, port)
        '    Try
        '        Do
        '            NetListener2.Start()
        '            Dim clib As New TcpClient
        '            clib = NetListener2.AcceptTcpClient
        '            Dim NetStream As NetworkStream = clib.GetStream
        '            Dim bytes(clib.ReceiveBufferSize) As Byte
        '            NetStream.Read(bytes, 0, CInt(clib.ReceiveBufferSize))
        '            Dim cliData As String = Encoding.ASCII.GetString(bytes)
        '            Msg = cliData + Chr(34)
        '            Dim msg2 As String
        '            msg2 = Mid(Msg, 1, 3)
        '            'MsgBox(msg2.ToString.Length)
        '            RaiseEvent GotMessage2(msg2)
        '            clib.Close()
        '            NetListener2.Stop()
        '        Loop Until False
        '    Catch ex As Exception
        '    End Try
        'Catch ex As Exception

        'End Try
    End Sub
    Public Sub DisplayNetMsg(ByVal MSG As String)

        'Dim x As New cmraude.getComBlockNFO
        'x.GetConNFO(Trim(MSG))
        'Dim response As Integer = 0

        'Dim z As New Transfer_Lead2
        'z.Text = "Transfer Lead: " & MSG.ToString
        'z.txtLeadNUM.Text = WhoC
        'z.txtLine.Text = LineC
        'z.rtfNotes.Text = NoteC
        'z.txtAR.Text = AutoC
        'z.Size = New Size(378, 305)

        'z.Timer1.Enabled = True
        'z.ShowInTaskbar = False
        'z.ShowDialog()

    End Sub
    Public Sub DisplayAcceptDecline(ByVal MSG As String)
        'Select Case MSG
        '    Case Is = "Acc"
        '        MsgBox("Client Accepted Transfer.", MsgBoxStyle.Information, "Transfer Accepted")
        '        Exit Select
        '    Case Is = "Dec"
        '        MsgBox("Client Declined Transfer", MsgBoxStyle.Information, "Transfer Declined")
        '        Exit Select
        '    Case Is = "Tim"
        '        MsgBox("Client did not respond in the alloted time.", MsgBoxStyle.Information, "Transfer Timed Out")
        '        Exit Select
        'End Select
    End Sub

    Private Sub ColdCallingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColdCallingToolStripMenuItem.Click

    End Sub

    Private Sub WarmCallingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WarmCallingToolStripMenuItem.Click
        Dim x As Form
        x = WCaller
        x.MdiParent = Me
        x.Show()
    End Sub



    Private Sub tsbfind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbfind.Click
        FindLead.PREVIOUS_ID = "0"
        FindLead.ShowInTaskbar = False
        FindLead.ShowDialog()
        If Me.forcettconfirm = True Then
            Confirming.Create_Tooltip(Confirming.dpConfirming, "search results")
        End If
        If Me.forcettsales = True Then
            Confirming.Create_Tooltip(Confirming.dpSales, "search results")
        End If


    End Sub

    Private Sub tsbchat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbchat.Click
        Dim x As Form
        x = frmChat
        'STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
        x.Show()
    End Sub
    Private Sub Send(ByVal t As String)
        Dim w As New IO.StreamWriter(mobjClient.GetStream)
        w.Write(t & vbCr)
        w.Flush()
    End Sub
    Private Sub DoRead(ByVal ar As IAsyncResult)
        '' do read takes the iasynresult as an argument
        '' the logic here is to just look to see if there is any data being sent in the underlying stream
        '' assuming that iaysncresult is returning true, there is a line underneath
        '' if the line data length is less than one, drop the line, there isnt one.
        '' from here, if the line length is > 1 AND iasyncresult is true,
        '' pass the message to stringbuilder object as an array of bytes (MARDATA)
        '' 
        'Dim intCount As Integer

        'Try
        '    intCount = mobjClient.GetStream.EndRead(ar)
        '    If intCount < 1 Then
        '        MarkAsDisconnected()
        '        Exit Sub
        '    End If

        '    BuildString(marData, 0, intCount)

        '    mobjClient.GetStream.BeginRead(marData, 0, 1024, AddressOf DoRead, Nothing)
        'Catch e As Exception
        '    MarkAsDisconnected()
        'End Try
    End Sub
    Private Sub BuildString(ByVal Bytes() As Byte, ByVal offset As Integer, ByVal count As Integer)
        '' bytes is the actual messages being sent in a byte array()
        '' offset is the marker of where to start inside of the byte array
        '' count is the length of the message.
        '' 
        '' the basic loop logic is intindex is the start position = where you should start , plus the count of the msg -1 character
        '' if the character in the position in the loop is a line feed character, append the line feed to the msg
        '' then reset the stringbuilder to be a new stringbuilder.
        '' append message until you get to a line feed character, then recycle the stringbuilder object.
        '' 
        'Dim intIndex As Integer

        'For intIndex = offset To offset + count - 1
        '    If Bytes(intIndex) = 10 Then
        '        mobjText.Append(vbLf)


        '        '' why use a displayinvoker here?
        '        '' invalidate the crossthreadcheck and just use a standard sub
        '        '' to write the message somewhere...

        '        Dim params() As Object = {mobjText.ToString}

        '        Me.Invoke(New DisplayInvoker(AddressOf Me.DisplayText), params)

        '        mobjText = New StringBuilder()
        '    Else
        '        mobjText.Append(ChrW(Bytes(intIndex)))
        '    End If
        'Next
    End Sub
    Private Sub MarkAsDisconnected()
        'MsgBox("Server Unavailable/Disconnected.")
        'txtSend.ReadOnly = True
        'btnSend.Enabled = False
    End Sub

    Private Sub DisplayText(ByVal t As String)
        'txtDisplay.AppendText(t)
        'MsgBox(t.ToString)
    End Sub

    Private Sub ConfirmingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfirmingToolStripMenuItem.Click
        Confirming.MdiParent = Me
        Confirming.Show()
        Confirming.BringToFront()


    End Sub

    Private Sub NewUserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SetUpUser.ShowInTaskbar = False
        SetUpUser.ShowDialog()

    End Sub

    Private Sub tsbmap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbmap.Click
        'System.Diagnostics.Process.Start("C:\Users\Andy\Documents\Programming Files\ISSMappointEXE\ISSMappointEXE\bin\Debug\ISSMappointEXE.exe")
        'MsgBox("Place holder for external mappoint.")
        frmMappoint.Show()
        frmMappoint.BringToFront()
        frmMappoint.MdiParent = Me

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrXFER.Tick
        If STATIC_VARIABLES.PendingXFER = True Then
            STATIC_VARIABLES.PendingXFER = False
            'Transfer_Lead2.Show()
        End If
    End Sub

    Private Sub tmrStartupLauncher_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStartupLauncher.Tick
        'Me.tmrStartupLauncher.Stop()
        'Me.tmrStartupLauncher.Dispose()
        'If PENDING_REQUEST_INFO.GetStartupForm = True Then

        '    '' need a way to ctype the startup form to an actual form and then show it.
        '    '' 
        '    '' no: select case the string to launch the appropriate form.

        '    Select Case STATIC_VARIABLES.StartUpForm
        '        Case Is = "Cold Calling"
        '            '' need form
        '            ColdCalling.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Warm Calling"
        '            WCaller.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Previous Customer"
        '            '' need form
        '            PreviousCustomer.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Recovery"
        '            '' need form
        '            Recovery.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Marketing Manager"
        '            '' need form
        '            MarketingManager.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Sales Department"
        '            '' need form
        '            Sales.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Installation"
        '            '' need form
        '            Installation.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Finance"
        '            '' need form
        '            Finance.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Administration"
        '            '' need form
        '            Administration.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Is = "Confirmer"
        '            Confirming.Show()
        '            PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '        Case Else
        '            'Main.Show()
        '            'PENDING_REQUEST_INFO.GetStartupForm = False
        '            Exit Select
        '    End Select



        'End If
    End Sub



    Private Sub BugsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BugsToolStripMenuItem.Click
        Dim day As String = Date.Now.Day
        Dim month As String = Date.Now.Month
        Dim year As String = Date.Now.Year
        'MsgBox(month & "/" & day & "/" & year)
        Dim suggest As String = InputBox$("Enter the bug you have found. Please be as descriptive as possible." & vbCr & "Example: On [form used] I was [doing this action] and [this] happened.", "Bug to Report", "", , )
        If suggest.ToString.Length <= 0 Then
            Exit Sub
        ElseIf suggest.ToString.Length >= 1 Then
            Dim str As New IO.StreamWriter("\\server\iss\company\Bugs\" & month & "-" & day & "-" & year & " BUGS.txt", True)
            str.WriteLine(STATIC_VARIABLES.CurrentUser.ToString & " | " & STATIC_VARIABLES.IP.ToString & " | " & suggest.ToString & vbCr)
            str.Flush()
            str.Close()
        End If
    End Sub

    Private Sub EnterASuggestionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnterASuggestionToolStripMenuItem.Click
        Dim day As String = Date.Now.Day
        Dim month As String = Date.Now.Month
        Dim year As String = Date.Now.Year
        'MsgBox(month & "/" & day & "/" & year)
        Dim suggest As String = InputBox$("Suggestions?", "Suggestion", "", , )
        If suggest.ToString.Length <= 0 Then
            Exit Sub
        ElseIf suggest.ToString.Length >= 1 Then
            Dim str As New IO.StreamWriter("\\server\iss\company\Suggestions\" & month & "-" & day & "-" & year & " suggestions.txt", True)
            str.WriteLine(STATIC_VARIABLES.CurrentUser.ToString & " | " & STATIC_VARIABLES.IP.ToString & " | " & suggest.ToString & vbCr)
            str.Flush()
            str.Close()
        End If
    End Sub

    Private Sub CompanyInformationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompanyInformationToolStripMenuItem.Click
        'frmEditCompanyInformation.MdiParent = Me
        frmEditCompanyInformation.StartPosition = FormStartPosition.CenterScreen
        frmEditCompanyInformation.ShowInTaskbar = False
        frmEditCompanyInformation.ShowDialog()

    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        frmViewEditAutoNotes.MdiParent = Me
        frmViewEditAutoNotes.StartPosition = FormStartPosition.CenterScreen
        frmViewEditAutoNotes.Show()
    End Sub

    Private Sub WorkHoursToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkHoursToolStripMenuItem.Click
        frmViewEditWorkHours.MdiParent = Me
        frmViewEditWorkHours.StartPosition = FormStartPosition.CenterScreen
        frmViewEditWorkHours.Show()
    End Sub

    Private Sub PrimaryLeadSourcesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrimaryLeadSourcesToolStripMenuItem.Click
        frmViewEditPrimaryLeadSources.MdiParent = Me
        frmViewEditPrimaryLeadSources.StartPosition = FormStartPosition.CenterScreen
        frmViewEditPrimaryLeadSources.Show()
    End Sub

    Private Sub ProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductsToolStripMenuItem.Click
        frmViewEditProducts.MdiParent = Me
        frmViewEditProducts.StartPosition = FormStartPosition.CenterScreen
        frmViewEditProducts.Show()
    End Sub

    Private Sub ErrorLogsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ErrorLogsToolStripMenuItem.Click
        frmErrorLogs.MdiParent = Me
        frmErrorLogs.StartPosition = FormStartPosition.CenterParent
        frmErrorLogs.Show()
    End Sub

    Private Sub BackupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackupToolStripMenuItem.Click
        frmBackup.MdiParent = Me
        frmBackup.StartPosition = FormStartPosition.CenterParent
        frmBackup.Show()
    End Sub


    Private Sub WindowsToolStripMenuItem_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles WindowsToolStripMenuItem.DropDownOpening
        Dim count As Integer = Me.MdiChildren.Length
        Me.WindowsToolStripMenuItem.DropDownItems.Clear()
        Me.WindowsToolStripMenuItem.DropDownItems.Add(CloseAll)
        CloseAll.Text = "Close All"

        Me.WindowsToolStripMenuItem.DropDownItems.Add(New ToolStripSeparator)
        Me.WindowsToolStripMenuItem.DropDownItems.Add(TileVertically)
        TileVertically.Text = "Tile Vertically"

        Me.WindowsToolStripMenuItem.DropDownItems.Add(TileHorizontally)
        TileHorizontally.Text = "Tile Horizontally"
        Me.WindowsToolStripMenuItem.DropDownItems.Add(Cascade)
        Cascade.Text = "Cascade"


        Me.WindowsToolStripMenuItem.DropDownItems.Add(New ToolStripSeparator)
        For i As Integer = 0 To count - 1
            frmbutton = New ToolStripMenuItem
            frmbutton.Text = Me.MdiChildren(i).Text
            frmbutton.Tag = i
            Dim c As New ContextMenu
            c.MenuItems.Add("Close " & Me.MdiChildren(i).Text)


            If Me.MdiChildren(i) Is Me.ActiveMdiChild Then
                'frmbutton.Image = Me.ILIcons.Images(0)
                frmbutton.CheckState = CheckState.Checked
            End If
            Me.WindowsToolStripMenuItem.DropDownItems.Add(frmbutton)
            AddHandler frmbutton.Click, AddressOf frmbutton_Click
        Next




        Me.WindowsToolStripMenuItem.DropDown.PerformLayout()
    End Sub

    Private Sub frmbutton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles frmbutton.Click
        Dim Rb As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Me.ActivateMdiChild(MdiChildren(CInt(Rb.Tag)))
        Me.MdiChildren(CInt(Rb.Tag)).Focus()
    End Sub




    Private Sub tmrmanagealerts_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrmanagealerts.Tick
        'Me.tmrmanagealerts.Stop()


        'Dim User As String = STATIC_VARIABLES.CurrentUser

        'Dim x As New ALERT_LOGIC(User)
        'Select Case x.CountOfAlerts
        '    Case Is >= 1

        '        Dim response As Integer = MsgBox("You have past due alerts. Would you like to manage them now?", MsgBoxStyle.YesNo, "Past Due Alerts")

        '        'Main.BringToFront()
        '        If response = 6 Then
        '            'ManageAlerts.MdiParent = Main
        '            'ManageAlerts.ShowInTaskbar = False
        '            'ManageAlerts.StartPosition = FormStartPosition.CenterParent
        '            'ManageAlerts.WindowState = FormWindowState.Normal
        '            If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
        '                STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
        '            End If
        '            ManageAlerts.Show()
        '            'ManageAlerts.TopMost = True

        '            'ManageAlerts.PerformLayout()
        '        ElseIf response = 7 Then
        '            Me.tmrAlerts.Start()
        '            ManageAlerts.Dispose()
        '        End If
        '    Case Is <= 0
        '        Me.tmrAlerts.Start()
        '        ManageAlerts.Dispose()
        '        Exit Select
        'End Select
        'Me.tmrmanagealerts.Dispose()

        'Confirming.WindowState = FormWindowState.Maximized
    End Sub

   

    Private Sub Cascade_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cascade.Click
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.Cascade)
    End Sub

    Private Sub TileHorizontally_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TileHorizontally.Click
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileHorizontal)
    End Sub

    Private Sub TileVertically_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TileVertically.Click
        Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileVertical)
    End Sub

    Private Sub CloseAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CloseAll.Click
        For Each frm As Form In Me.MdiChildren
            frm.Close()
        Next
        Me.Refresh()
    End Sub

    Private Sub Main_MdiChildActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MdiChildActivate

        Dim count As Integer = Me.MdiChildren.Length
        For i As Integer = 0 To count - 1
            'Me.MdiChildren(i).WindowState = FormWindowState.Normal
            If Me.MdiChildren(i) Is Me.ActiveMdiChild Then
                Dim f As Form = Me.MdiChildren(i)
                Select Case f.Name
                    Case "Confirming", "Sales", "WCaller", "MarketingManager", "Finance", "Installation", "Administration", _
                    "Recovery", "PreviousCustomer", "ColdCalling", "ConfirmingSingleRecord", "SecondSource"
                        STATIC_VARIABLES.ActiveChild = Me.MdiChildren(i)
                        ''Add decision structure here to manage current id 
                        '' may fire this event before it selects listview item could be problamatic ''It wasnt 
                        '' hopefully this fires after load event on form 
                        ''Tested, Load event for form happens then it hit mdiactivate event here, Should work perfect,
                        '' and eliminate the need to manage form by form, but it will still be neccesary to cahnge currentID on tab change in each form
                        Dim lv As ListView
                        Select Case f.Name
                            Case "Confirming"
                                If Confirming.TabControl1.SelectedIndex = 0 Then
                                    lv = Confirming.lvConfirming
                                ElseIf Confirming.TabControl1.SelectedIndex = 1 Then
                                    lv = Confirming.lvSales
                                Else
                                    lv = Nothing
                                End If
                            Case "Sales"
                                If Sales.tbMain.SelectedIndex = 0 Then
                                    If Sales.lvnoresults.SelectedItems.Count <> 0 Then
                                        lv = Sales.lvnoresults
                                    Else
                                        lv = Nothing
                                    End If
                                ElseIf Sales.tbMain.SelectedIndex = 1 Then
                                    If Sales.TabControl2.SelectedIndex = 0 Then
                                        lv = Sales.lvSales
                                    ElseIf Sales.TabControl2.SelectedIndex = 1 Then
                                        lv = Nothing
                                    Else
                                        lv = Nothing
                                    End If
                                End If

                            Case "Wcaller"
                                If WCaller.TabControl2.SelectedIndex = 0 Then
                                    lv = WCaller.lvWarmCalling
                                ElseIf WCaller.TabControl2.SelectedIndex = 1 Then
                                    lv = WCaller.lvMyAppts
                                Else
                                    lv = Nothing
                                End If
                            Case "MarketingManager"
                                ''add case as form complete 
                            Case "Finance"
                                ''add case as form complete
                            Case "Installation"
                                ''add case as form complete
                            Case "Administration"
                                ''add case as form complete
                            Case "Recovery"
                                ''add case as form complete
                            Case "PreviousCustomer"
                                ''add case as form complete
                            Case "ColdCalling"
                                ''add case as form complete
                            Case "ConfirmingSingleRecord"
                                lv = Nothing
                                STATIC_VARIABLES.CurrentID = ConfirmingSingleRecord.ID
                            Case "SecondSource"
                        End Select
                        If lv IsNot Nothing Then
                            If lv.SelectedItems.Count <> 0 And lv IsNot Sales.lvnoresults Then
                                STATIC_VARIABLES.CurrentID = lv.SelectedItems(0).Text
                            ElseIf lv.SelectedItems.Count <> 0 And lv Is Sales.lvnoresults Then
                                STATIC_VARIABLES.CurrentID = lv.SelectedItems(0).SubItems(1).Text
                            Else
                                STATIC_VARIABLES.CurrentID = ""
                            End If

                        Else
                            If f.Name = "ConfirmingSingleRecord" Then
                                STATIC_VARIABLES.CurrentID = ConfirmingSingleRecord.ID
                            ElseIf f.Name = "Sales" And Sales.tbMain.SelectedIndex = 0 Then
                                If Sales.lvnoresults.SelectedItems.Count <> 0 Then
                                    STATIC_VARIABLES.CurrentID = Sales.lvnoresults.Tag
                                End If
                            ElseIf f.Name = "Sales" And Sales.tbMain.SelectedIndex = 1 And Sales.TabControl2.SelectedIndex = 1 Then
                                STATIC_VARIABLES.CurrentID = Sales.lvMemorized.SelectedItems.Item(0).SubItems.Item(1).Text
                            Else

                                STATIC_VARIABLES.CurrentID = ""
                            End If
                        End If
                    Case Else
                        'STATIC_VARIABLES.ActiveChild = Nothing
                        'STATIC_VARIABLES.CurrentID = ""
                End Select
            End If

        Next


    End Sub




    Dim fontsize As Single = 9.75



    Public Sub open()

        SetAppt.ID = Sales.ID
        SetAppt.frm = Sales
        SetAppt.OrigApptDate = Sales.txtApptDate.Text
        SetAppt.OrigApptTime = Sales.txtApptTime.Text
        Dim s = Split(Sales.txtContact1.Text, " ")
        Dim s2 = Split(Sales.txtContact2.Text, " ")
        SetAppt.Contact1 = s(0)
        SetAppt.Contact2 = s2(0)


        SetAppt.ShowDialog()
    End Sub


    Private Sub MarketingManagerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MarketingManagerToolStripMenuItem.Click

        MarketingManager.MdiParent = Me
        MarketingManager.Show()
        MarketingManager.BringToFront()

    End Sub

    Private Sub tsbadmin_Click(sender As Object, e As EventArgs) Handles tsbadmin.Click
        Administration.MdiParent = Me
        Administration.Show()
    End Sub

    Private Sub ImportDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportDataToolStripMenuItem.Click

    End Sub

    Private Sub Improveit360ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Improveit360ToolStripMenuItem.Click
        Form6.ShowDialog()
    End Sub

 
End Class
