

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

    Private mobjClient As TcpClient
    Private marData(1024) As Byte
    Private mobjText As New StringBuilder()
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
        x = sales
        sales.LOAD_HANDLER = "LOAD"
        x.MdiParent = Me
        x.Show()
    End Sub

    Private Sub tsbnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbnew.Click
        '' put this back to show enterlead inside of mdiparent
        '' 'Main'

        'Dim y As New Form
        'y = My.Forms.EnterLead

        EnterLead.MdiParent = Me
        EnterLead.Show()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub tsbattach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbattach.Click
        AttachAFile.Show()
        Try
            Dim y As String = InputBox$("Please enter the lead number you wish to attach the file(s) to.", "Enter Lead Number", STATIC_VARIABLES.CurrentID.ToString) ' RecordLogic.CurrentID was taken out for default recordID in rev 5
            If y = "" Then
                'MsgBox("You must supply a lead number to attach a file to.", MsgBoxStyle.Exclamation, "Error Attaching File")
                Exit Sub
            End If
            Dim z As New Attach
            Dim a As String = "0"
            z.AttachFile(y, a)
            'MsgBox("File attached to Lead#: " & y.ToString & "", MsgBoxStyle.Information, "File Attached Successfully")
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub tsImportsPics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsImportsPics.Click
        Dim x As Form
        x = ImportPictures
        x.MdiParent = Me
        x.Show()
    End Sub

    Private Sub tsbrolodex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbrolodex.Click
        Dim x As Form
        x = frmRolodex 'Employee_Contacts
        x.MdiParent = Me
        x.Show()
    End Sub

    Private Sub tsbtransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtransfer.Click
        Dim x As Form
        x = Transfer_Lead
        x.ShowDialog()
    End Sub

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        NetListener.Stop() '' caller listener shutdown
        NetListener2.Stop() '' accept/decline listener shutdown
        Dim c As New USER_LOGIC
        c.LogOut_Of_System(STATIC_VARIABLES.Login, STATIC_VARIABLES.UserPWD)
    End Sub

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'ThNet.Start()
        'ThAD.Start()

        Me.TSMain.Enabled = False

        tmrAlerts.Enabled = True
        tmrXFER.Enabled = True
        tmrStartupLauncher.Enabled = True

        tmrAlerts.Start()


        LOGIN.ShowDialog()
        Dim r As New REG_LOGIC
        r.ReadKey(STATIC_VARIABLES.LicenseKey, STATIC_VARIABLES.LeaseKey)
    End Sub
    Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        Dim r As New REG_LOGIC
        LOGIN.ShowDialog()
        r.ReadKey(STATIC_VARIABLES.LicenseKey, STATIC_VARIABLES.LeaseKey)
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Dim c As New USER_LOGIC
        c.LogOut_Of_System(STATIC_VARIABLES.CurrentUser, STATIC_VARIABLES.UserPWD)

    End Sub

    Private Sub tsbschedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbschedule.Click
        Dim x As Form
        x = ScheduleAction
        x.Show()
        x.Focus()
        x.TopMost = True

    End Sub

    Private Sub tsbalert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbalert.Click
        Dim x As Form
        x = AssignAlert
        x.MdiParent = Me
        x.Show()
    End Sub

    Private Sub tmrAlerts_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrAlerts.Tick
        tmrAlerts.Stop()
        Dim x As New ALERT_LOGIC(STATIC_VARIABLES.CurrentUser)
        Dim hour As Integer = Now.Hour
        Dim minute As Integer = Now.Minute
        Dim seconds As Integer = Now.Second
        Dim correctedTime As String = Date.Today.Date & " " & hour & ":" & minute & ":" & seconds
        Select Case x.CountOfAlerts
            Case Is >= 1
                ManageAlerts.ShowDialog()
                Exit Select
            Case Is <= 0
                Exit Select
        End Select
    End Sub
    Private Sub DoListen()
        Try
            AddHandler GotMessage, AddressOf DisplayNetMsg
            Control.CheckForIllegalCrossThreadCalls = False
            Dim port As Int32 = 5586 '' currently unassigned by IANA as of 2-17-2007
            Dim ip As IPAddress '= Me.GetIPv4Address.ToString                                        ' holder for local machine name.
            Dim strMachineName As String = ""
            strMachineName = My.Computer.Name
            Dim ipHost As IPHostEntry
            ipHost = Dns.GetHostEntry(strMachineName)
            Dim ipAddr() As IPAddress = ipHost.AddressList
            ip = System.Net.IPAddress.Any ''ip = GetIPv4Address()
            Dim ep As IPEndPoint = New IPEndPoint(ip, port)
            NetListener = New TcpListener(ep)
            Try
                'Do
                NetListener.Start()
                Dim clib As New TcpClient
                clib = NetListener.AcceptTcpClient
                Dim NetStream As NetworkStream = clib.GetStream
                Dim bytes(clib.ReceiveBufferSize) As Byte
                NetStream.Read(bytes, 0, CInt(clib.ReceiveBufferSize))
                Dim cliData As String = Encoding.ASCII.GetString(bytes)
                Msg = cliData.ToString
                Dim machine As String = "AndyGlynn2"
                Dim Ip2 As String = "192.168.1.8"
                Dim msg2 As String = "Go Fuck Yourself"
                Dim CMD As String = "LOAD"
                Dim USR As String = "AndyG"
                Dim Rec As String = "Server"











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
                RaiseEvent GotMessage(LeadNumC)
                clib.Close()
                NetListener.Stop()
                ' Loop Until False
            Catch ex As Exception
                MsgBox(ex.Message, , )

            End Try
        Catch ex As Exception
            MsgBox(ex.Message, , )
        End Try
    End Sub
    Private Sub DoListen2()
        Try
            AddHandler GotMessage2, AddressOf DisplayAcceptDecline
            Control.CheckForIllegalCrossThreadCalls = False
            Dim port As Int32 = 5587 '' currently unassigned by IANA as of 2-17-2007
            Dim ip As IPAddress ' holder for local machine name.
            Dim strMachineName As String = ""
            strMachineName = My.Computer.Name
            Dim ipHost As IPHostEntry
            ipHost = Dns.GetHostByName(strMachineName)
            Dim ipAddr() As IPAddress = ipHost.AddressList
            ip = GetIPv4Address()
            NetListener2 = New TcpListener(ip, port)
            Try
                Do
                    NetListener2.Start()
                    Dim clib As New TcpClient
                    clib = NetListener2.AcceptTcpClient
                    Dim NetStream As NetworkStream = clib.GetStream
                    Dim bytes(clib.ReceiveBufferSize) As Byte
                    NetStream.Read(bytes, 0, CInt(clib.ReceiveBufferSize))
                    Dim cliData As String = Encoding.ASCII.GetString(bytes)
                    Msg = cliData + Chr(34)
                    Dim msg2 As String
                    msg2 = Mid(Msg, 1, 3)
                    'MsgBox(msg2.ToString.Length)
                    RaiseEvent GotMessage2(msg2)
                    clib.Close()
                    NetListener2.Stop()
                Loop Until False
            Catch ex As Exception
            End Try
        Catch ex As Exception

        End Try
    End Sub
    Public Sub DisplayNetMsg(ByVal MSG As String)

        'Dim x As New cmraude.getComBlockNFO
        'x.GetConNFO(Trim(MSG))
        Dim response As Integer = 0

        Dim z As New Transfer_Lead2
        z.Text = "Transfer Lead: " & MSG.ToString
        z.txtLeadNUM.Text = WhoC
        z.txtLine.Text = LineC
        z.rtfNotes.Text = NoteC
        z.txtAR.Text = AutoC
        z.Size = New Size(378, 305)

        z.Timer1.Enabled = True
        z.ShowDialog()

    End Sub
    Public Sub DisplayAcceptDecline(ByVal MSG As String)
        Select Case MSG
            Case Is = "Acc"
                MsgBox("Client Accepted Transfer.", MsgBoxStyle.Information, "Transfer Accepted")
                Exit Select
            Case Is = "Dec"
                MsgBox("Client Declined Transfer", MsgBoxStyle.Information, "Transfer Declined")
                Exit Select
            Case Is = "Tim"
                MsgBox("Client did not respond in the alloted time.", MsgBoxStyle.Information, "Transfer Timed Out")
                Exit Select
        End Select
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
        FindLead.ShowDialog()
    End Sub

    Private Sub tsbchat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbchat.Click
        Dim x As Form
        x = frmChat
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
        Dim intCount As Integer

        Try
            intCount = mobjClient.GetStream.EndRead(ar)
            If intCount < 1 Then
                MarkAsDisconnected()
                Exit Sub
            End If

            BuildString(marData, 0, intCount)

            mobjClient.GetStream.BeginRead(marData, 0, 1024, AddressOf DoRead, Nothing)
        Catch e As Exception
            MarkAsDisconnected()
        End Try
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
        Dim intIndex As Integer

        For intIndex = offset To offset + count - 1
            If Bytes(intIndex) = 10 Then
                mobjText.Append(vbLf)


                '' why use a displayinvoker here?
                '' invalidate the crossthreadcheck and just use a standard sub
                '' to write the message somewhere...

                Dim params() As Object = {mobjText.ToString}

                Me.Invoke(New DisplayInvoker(AddressOf Me.DisplayText), params)

                mobjText = New StringBuilder()
            Else
                mobjText.Append(ChrW(Bytes(intIndex)))
            End If
        Next
    End Sub
    Private Sub MarkAsDisconnected()
        MsgBox("Server Unavailable/Disconnected.")
        'txtSend.ReadOnly = True
        'btnSend.Enabled = False
    End Sub

    Private Sub DisplayText(ByVal t As String)
        'txtDisplay.AppendText(t)
        MsgBox(t.ToString)
    End Sub
    
    Private Sub ConfirmingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfirmingToolStripMenuItem.Click
        Confirming.Show()

    End Sub

    Private Sub NewUserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewUserToolStripMenuItem.Click
        SetUpUser.ShowDialog()

    End Sub

    Private Sub tsbmap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbmap.Click
        System.Diagnostics.Process.Start("C:\Documents and Settings\xxclayxx\Desktop\ISSMappointEXE\ISSMappointEXE\bin\Debug\ISSMappointEXE.exe")
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrXFER.Tick
        If STATIC_VARIABLES.PendingXFER = True Then
            STATIC_VARIABLES.PendingXFER = False
            Transfer_Lead2.Show()
        End If
    End Sub

    Private Sub tmrStartupLauncher_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStartupLauncher.Tick
        If PENDING_REQUEST_INFO.GetStartupForm = True Then

            '' need a way to ctype the startup form to an actual form and then show it.
            '' 
            '' no: select case the string to launch the appropriate form.

            Select Case STATIC_VARIABLES.StartUpForm
                Case Is = "Cold Calling"
                    '' need form
                    ColdCalling.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Warm Calling"
                    WCaller.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Previous Customer"
                    '' need form
                    PreviousCustomer.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Recovery"
                    '' need form
                    Recovery.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Marketing Manager"
                    '' need form
                    MarketingManager.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Sales Department"
                    '' need form
                    SalesDepartment.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Installation"
                    '' need form
                    Installation.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Finance"
                    '' need form
                    Finance.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Administration"
                    '' need form
                    Administration.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Is = "Confirmer"
                    Confirming.Show()
                    PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
                Case Else
                    'Main.Show()
                    'PENDING_REQUEST_INFO.GetStartupForm = False
                    Exit Select
            End Select



        End If
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
            Dim str As New IO.StreamWriter("\\ekg1\iss\Bugs\" & month & "-" & day & "-" & year & " BUGS.txt", True)
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
            Dim str As New IO.StreamWriter("\\ekg1\iss\Suggestions\" & month & "-" & day & "-" & year & " suggestions.txt", True)
            str.WriteLine(STATIC_VARIABLES.CurrentUser.ToString & " | " & STATIC_VARIABLES.IP.ToString & " | " & suggest.ToString & vbCr)
            str.Flush()
            str.Close()
        End If
    End Sub

    Private Sub CompanyInformationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompanyInformationToolStripMenuItem.Click
        frmEditCompanyInformation.MdiParent = Me
        frmEditCompanyInformation.StartPosition = FormStartPosition.CenterScreen
        frmEditCompanyInformation.Show()

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
End Class
