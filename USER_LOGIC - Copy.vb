
Imports System.Data
Imports System
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Net.Sockets
Imports System.Net
Imports System.Text
Imports System.IO


Public Class USER_LOGIC
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Public mardata(1024) As Byte
    Public mobjClient As TcpClient
    Private mobjText As New StringBuilder()
    Public Delegate Sub DisplayInvoker(ByVal t As String)

    Public Function GetIPv4Address() As IPAddress
        GetIPv4Address = Nothing
        Dim strmachine As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strmachine)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPv4Address = ipheal
                ' MsgBox(ipheal.ToString, MsgBoxStyle.Critical, "ERROR")

            End If
        Next

    End Function ''Good
    Public Sub Login_To_System()

        Dim USR As String = ""
        Dim USRF As String = ""
        Dim USRL As String = ""
        Dim PWD As String = ""
        If LOGIN.txtUserName.Text = "Admin" Then
            USR = "Admin"
            ' USRL = "Admin"
            PWD = LOGIN.txtPWD.Text
        End If
        If LOGIN.txtUserName.Text <> "Admin" Then
            ' Dim str2 = Split(LOGIN.txtUserName.Text, " ", 2)

            ' USR = Trim(str2(0))
            ' USRL = Trim(str2(1))
            USR = LOGIN.txtUserName.Text
            PWD = LOGIN.txtPWD.Text
        End If


        If USR.ToString.Length < 2 Then
            MsgBox("User Name cannot be blank.", MsgBoxStyle.Critical, "ERROR")
            Exit Sub
        End If

        If PWD.ToString.Length < 2 Then
            MsgBox("Password cannot be blank.", MsgBoxStyle.Critical, "ERROR")
            Exit Sub
        End If

        Dim cmdCNT As SqlCommand = New SqlCommand("SELECT Count(ID) from iss.dbo.userpermissiontable where UserName = @USR and UserPWD = @PWD", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@USR", USR)
        'Dim param33 As SqlParameter = New SqlParameter("@USRL", USRL)
        Dim param2 As SqlParameter = New SqlParameter("@PWD", PWD)
        cmdCNT.Parameters.Add(param1)
        cmdCNT.Parameters.Add(param2)
        '  cmdCNT.Parameters.Add(param33)
        cnn.Open()
        Dim r1 As SqlDataReader
        Dim cnt As Integer = 0
        r1 = cmdCNT.ExecuteReader
        While r1.Read
            cnt = r1.Item(0)
        End While
        r1.Close()
        cnn.Close()
        Try
            Select Case cnt
                Case Is <= 0
                    MsgBox("User does not exist. Please try again.", MsgBoxStyle.Critical, "ERROR")
                    Exit Sub
                Case Is = 1
                    '' pipe information from sql table to user logic module to retain values all through out the program
                    '' multiple checks.

                    Dim cmdGet As SqlCommand = New SqlCommand("dbo.Get_User_Info", cnn)
                    cmdGet.CommandType = CommandType.StoredProcedure
                    Dim param22 As SqlParameter = New SqlParameter("@UserName", USR)
                    'Dim param3 As SqlParameter = New SqlParameter("@UserFirstName", USRF)
                    'Dim param44 As SqlParameter = New SqlParameter("@UserLastName", USRL)
                    Dim param4 As SqlParameter = New SqlParameter("@UserPWD", PWD)
                    cmdGet.Parameters.Add(param22)
                    'cmdGet.Parameters.Add(param3)
                    cmdGet.Parameters.Add(param4)
                    'cmdGet.Parameters.Add(param44)

                    cnn.Open()
                    Dim r2 As SqlDataReader
                    r2 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                    While r2.Read
                        If USR = "Admin" Then
                            STATIC_VARIABLES.CurrentUser = "Admin"
                        ElseIf USR <> "Admin" Then
                            STATIC_VARIABLES.CurrentUser = r2.Item(2) & " " & r2.Item(3)
                        End If
                        STATIC_VARIABLES.Login = r2.Item(1)
                        STATIC_VARIABLES.CurrentUser = r2.Item(2) & " " & r2.Item(3)
                        STATIC_VARIABLES.UserPWD = r2.Item(4)
                        STATIC_VARIABLES.ColdCall = r2.Item(5)
                        STATIC_VARIABLES.WarmCall = r2.Item(6)
                        STATIC_VARIABLES.PreviousCust = r2.Item(7)
                        STATIC_VARIABLES.Recovery = r2.Item(8)
                        STATIC_VARIABLES.Confirmer = r2.Item(9)
                        STATIC_VARIABLES.SalesManager = r2.Item(10)
                        STATIC_VARIABLES.MarketingManager = r2.Item(11)
                        STATIC_VARIABLES.Finance = r2.Item(12)
                        STATIC_VARIABLES.Install = r2.Item(13)
                        STATIC_VARIABLES.Administration = r2.Item(14)
                        '' division wich was nixed = item 13
                        STATIC_VARIABLES.StartUpForm = r2.Item(15)
                        STATIC_VARIABLES.CurrentForm = r2.Item(15) '' on startup this variable needs a default value // will be replaced as soon a child form is opened up. but, for alerts and past due alerts need default value.
                        STATIC_VARIABLES.LoggedOn = r2.Item(16)
                        STATIC_VARIABLES.DoNotShowMapping = r2.Item(17)
                        STATIC_VARIABLES.MachineName = r2.Item(18)
                        STATIC_VARIABLES.ManagerFirstName = r2.Item(19)
                        STATIC_VARIABLES.ManagerLastName = r2.Item(20)
                        STATIC_VARIABLES.IP = r2.Item(21)
                        STATIC_VARIABLES.LicenseKey = r2.Item(22)
                        STATIC_VARIABLES.LeaseKey = r2.Item(23)
                    End While
                    r2.Close()
                    cnn.Close()



                    '' get use machine name and ip and log to table
                    ''

                    Dim strMachine As String = ""
                    strMachine = My.Computer.Name
                    'strMachine = STATIC_VARIABLES.MachineName
                    Dim ipHost As IPHostEntry
                    ipHost = Dns.GetHostEntry(strMachine)
                    Dim ip As IPAddress ' holder for local machine name.
                    Dim ipAddr() As IPAddress = Dns.GetHostAddresses(strMachine)

                    ip = ipAddr(0)
                    'Dim SpIP As String = ""
                    ''SpIP = Split(ipAddr(0).ToString, ".", 4)
                    'Dim ConvertedIP As String = ""
                    'ConvertedIP = SpIP(0) & "." & SpIP(1) & "." & SpIP(2) & "." & SpIP(3)
                    Dim ConvertedIP As String = Me.GetIPv4Address.ToString
                    STATIC_VARIABLES.IP = ConvertedIP
                    STATIC_VARIABLES.MachineName = strMachine
                 
                    ' MsgBox("Connected to server..." & STATIC_VARIABLES.CurrentUser & STATIC_VARIABLES.MachineName.ToString & STATIC_VARIABLES.IP.ToString & "WTF?", MsgBoxStyle.Information)





                    Try
                        '' reactivation code
                        '' uncomment
                        'Dim ipHost1 As IPHostEntry
                        'ipHost = Dns.GetHostEntry(Dns.GetHostName())
                        'Dim g
                        'g = ipHost.AddressList()

                        ' server name is statically assigned here.
                        ' must make it dynamic after install to pick up on the server name.
                        ' IE XML config file or something. No ideas yet.
                        ' SERVERS: 'EKG1' | 'DELLXPS'

                        ' need to shut server down right now to get this to work.
                        ' MUST re-activate at later time.

                        '  Dim ep As IPEndPoint = New IPEndPoint(, 5586)

                        mobjClient = New TcpClient(STATIC_VARIABLES.MachineName, 5586)
                        'MsgBox("Connected to server..." & STATIC_VARIABLES.MachineName, MsgBoxStyle.Information)

                        mobjClient.GetStream.BeginRead(mardata, 0, 1024, AddressOf DoRead, Nothing)



                        'Dim port As Int32 = 5586 '' currently unassigned by IANA as of 2-17-2007
                        'Dim ip1 As IPAddress ' holder for local machine name.
                        ''Dim strMachineName As String = ""
                        ''strMachineName = My.Computer.Name
                        'Dim ipHost As IPHostEntry
                        'ipHost1 = Dns.GetHostEntry(STATIC_VARIABLES.MachineName)
                        'Dim ipAddr1() As IPAddress = ipHost.AddressList
                        'ip1 = ipAddr1(0)

                        'Dim dnsEntry As String = ""
                        'dnsEntry = Dns.GetHostEntry(IP).HostName

                        ' packet structure 
                        ' machinename|ipaddress|msg|cmd|user|recipient (if any)|"

                        Send(PENDING_REQUEST_INFO.LeadNumber & "|" & STATIC_VARIABLES.MachineName & "|" & STATIC_VARIABLES.IP & "| |" & "LOAD" & "|" & STATIC_VARIABLES.Login & "| |")

                        STATIC_VARIABLES.NET_CLIENT = mobjClient
                        'MsgBox(STATIC_VARIABLES.NET_CLIENT.ToString)

                        ' no use a store procedure to write this information to the table
                        ' at login.

                        Dim cmdNET As SqlCommand = New SqlCommand("dbo.InsertNetworkINFO", cnn)
                        Dim param11 As SqlParameter
                        'Dim param55 As SqlParameter
                        If STATIC_VARIABLES.CurrentUser = "Admin" Then
                            param11 = New SqlParameter("@UserName", STATIC_VARIABLES.Login.ToString)
                            'param55 = New SqlParameter("@UserLName", STATIC_VARIABLES.CurrentUser.ToString)
                        ElseIf STATIC_VARIABLES.CurrentUser <> "Admin" Then
                            'Dim str = Split(STATIC_VARIABLES.CurrentUser, " ", 2)
                            param11 = New SqlParameter("@UserName", STATIC_VARIABLES.Login.ToString)
                            ' param55 = New SqlParameter("@UserLName", str(1))
                        End If
                        Dim param12 As SqlParameter = New SqlParameter("@UserPWD", STATIC_VARIABLES.UserPWD)
                        Dim param13 As SqlParameter = New SqlParameter("@MachineName", strMachine)
                        Dim param14 As SqlParameter = New SqlParameter("@IP", ConvertedIP)
                        'STATIC_VARIABLES.IP = ConvertedIP.ToString
                        cmdNET.CommandType = CommandType.StoredProcedure
                        cnn.Open()
                        cmdNET.Parameters.Add(param11)
                        cmdNET.Parameters.Add(param12)
                        cmdNET.Parameters.Add(param13)
                        cmdNET.Parameters.Add(param14)
                        'cmdNET.Parameters.Add(param55)
                        Dim r3 As SqlDataReader
                        r3 = cmdNET.ExecuteReader(CommandBehavior.CloseConnection)
                        r3.Close()
                        cnn.Close()


                        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE iss.dbo.userpermissiontable " _
                               & " SET LoggedOn = 1 " _
                               & " WHERE UserName = @USR  and UserPWD = @PWD", cnn)
                        Dim param78 As SqlParameter = New SqlParameter("@USR", USR)
                        'Dim param80 As SqlParameter = New SqlParameter("@USRL", USRL)
                        Dim param79 As SqlParameter = New SqlParameter("@PWD", UserPWD)
                        cmdUP.Parameters.Add(param78)
                        cmdUP.Parameters.Add(param79)
                        'cmdUP.Parameters.Add(param80)
                        cnn.Open()
                        Dim r4 As SqlDataReader
                        r4 = cmdUP.ExecuteReader(CommandBehavior.CloseConnection)
                        r4.Close()
                        cnn.Close()


                        ' MsgBox(STATIC_VARIABLES.Server_Assigned_Hash.ToString & ", " & STATIC_VARIABLES.Login.ToString & "," & STATIC_VARIABLES.CurrentUser.ToString & ", " & STATIC_VARIABLES.CurrentForm.ToString & ", " & STATIC_VARIABLES.MachineName & ", " & STATIC_VARIABLES.IP & ", " & STATIC_VARIABLES.StartUpForm)


                        '' now startup default form
                        '' 

                        Select Case STATIC_VARIABLES.StartUpForm
                            Case Is = "Cold Calling"
                                '' need form
                                ColdCalling.Show()
                                Exit Select
                            Case Is = "Warm Calling"
                                WCaller.Show()
                                Exit Select
                            Case Is = "Previous Customer"
                                '' need form
                                PreviousCustomer.Show()
                                Exit Select
                            Case Is = "Recovery"
                                '' need form
                                Recovery.Show()
                                Exit Select
                            Case Is = "Marketing Manager"
                                '' need form
                                MarketingManager.Show()
                                Exit Select
                            Case Is = "Sales Department"
                                '' need form
                                SalesDepartment.Show()
                                Exit Select
                            Case Is = "Installation"
                                '' need form
                                Installation.Show()
                                Exit Select
                            Case Is = "Finance"
                                '' need form
                                Finance.Show()
                                Exit Select
                            Case Is = "Administration"
                                '' need form
                                Administration.Show()
                                Exit Select
                            Case Is = "Confirmer"
                                Confirming.MdiParent = Main
                                Confirming.Show()
                                Confirming.BringToFront()
                                Confirming.Focus()
                                Confirming.Show()
                                Exit Select
                            Case Else
                                Main.Show()
                                Exit Select
                        End Select


                        '' now set the hash for xfer objects
                        ''


                        '' now check permissions and lock down appropriate buttons on tsmain
                        ''


                        CheckPermissions()


                        Try
                            Dim cmdINS As SqlCommand = New SqlCommand("INSERT iss.dbo.tblUserHash (UserName,UserHASH) " _
                            & "VALUES(@USR,@USRL,@HASH)", cnn)
                            Dim param100 As SqlParameter = New SqlParameter("@USR", Trim(STATIC_VARIABLES.Login.ToString))
                            ' Dim param200 As SqlParameter = New SqlParameter("@USRL", Trim(STATIC_VARIABLES.CurrentUser))
                            Dim param300 As SqlParameter = New SqlParameter("@HASH", Trim(STATIC_VARIABLES.Server_Assigned_Hash))
                            cmdINS.Parameters.Add(param100)
                            'cmdINS.Parameters.Add(param200)
                            cmdINS.Parameters.Add(param300)
                            cnn.Open()
                            cmdINS.ExecuteNonQuery()
                            cnn.Close()
                        Catch ex As Exception
                            cnn.Close()
                            MsgBox("Problem inserting hash to table.", MsgBoxStyle.Critical, "ERROR")
                        End Try

                    Catch ex As Exception
                        MsgBox("ip host ln 303", MsgBoxStyle.Critical, "ERROR")
                    End Try

                    '' now reset the 'Lead Pool' of 'Set Appointments'
                    ''

                    Try
                        Dim cmdRES As SqlCommand = New SqlCommand("dbo.ResetLeadPool", cnn)
                        cmdRES.CommandType = CommandType.StoredProcedure
                        cnn.Open()
                        cmdRES.ExecuteNonQuery()
                        cnn.Close()
                    Catch ex As Exception
                        MsgBox("Reset lead pool", MsgBoxStyle.Critical, "ERROR")
                        cnn.Close()
                    End Try

                    Exit Select

                Case Else
                    MsgBox("There was a problem logging in. Please contact your administrator.", MsgBoxStyle.Critical, "ERROR")
                    Exit Sub


            End Select
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("USER_LOGIC", "None", ex.Message.ToString, "Client", USR & " " & USRL & ", Login", "SQL", "Login_To_System")
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, "ERROR")

        End Try

    End Sub
    Private Sub Send(ByVal t As String)
        Try
            Dim w As New IO.StreamWriter(mobjClient.GetStream)
            w.Write(t & vbCr)
            'MsgBox(t)
            w.Flush()

        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("USER_LOGIC", "ByVal t as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Network", "Send")

        End Try


    End Sub
    Private Sub BuildString(ByVal Bytes() As Byte, ByVal offset As Integer, ByVal count As Integer)
        Try
            Dim intIndex As Integer

            For intIndex = offset To offset + count - 1
                If Bytes(intIndex) = 10 Then
                    mobjText.Append(vbLf)

                    Dim params() As Object = {mobjText.ToString}

                    Dim z As New DisplayInvoker(AddressOf Me.DisplayText)

                    mobjText = New StringBuilder()
                Else
                    mobjText.Append(ChrW(Bytes(intIndex)))

                End If
            Next
            Dim str
            str = Split(mobjText.ToString, "|", 7)
            Dim Mach As String = ""
            Dim IP As String = ""
            Dim MSG As String = ""
            Dim CMD As String = ""
            Dim USR As String = ""
            Dim Recip As String = ""
            Mach = Trim(str(0))
            IP = Trim(str(1))
            MSG = Trim(str(2))
            CMD = Trim(str(3))
            USR = Trim(str(4))
            Recip = Trim(str(5))
            Select Case USR
                Case Is = "SRV"
                    Select Case CMD
                        Case Is = "RCVHASH"
                            STATIC_VARIABLES.Server_Assigned_Hash = Trim(MSG.ToString)

                            MsgBox(STATIC_VARIABLES.Server_Assigned_Hash.ToString & " LOGGED IN")
                            mobjText = New StringBuilder()
                            Exit Select
                        Case Is = "ANNOUNCE"
                            'MsgBox(MSG & " online.")
                            mobjText = New StringBuilder()
                            Exit Select
                        Case Is = "XFER"
                            '' msg object needs to be split up
                            '' 
                            '' will need to be accurrately split apart here.
                            '' basically MACH|IP|LINE:LEADNUM:WHO:RESPONSE:NOTES|CMD|USR|RECIPIENT
                            ''                    0     1      2    3        4 

                            Dim strX = Split(MSG, ":", -1)
                            Dim line As String = ""
                            Dim leadnum As String = ""
                            Dim who As String = ""
                            Dim response As String = ""
                            Dim notes As String = ""

                            line = strX(0).ToString
                            leadnum = strX(1).ToString
                            who = strX(2).ToString
                            response = strX(3).ToString
                            notes = strX(4).ToString

                            'MsgBox(Recip & " is sending an xfer request on line: " & line.ToString & " for Lead#: " & leadnum & vbCr & " with Notes: " & notes.ToString)

                            '' notes had to swith over to another module to retain the values, then 
                            '' used a secondary timer to look for a pending xfer flag.
                            '' there is apparently an issue with the .net framework 2.0 and using asynchronous sockets.
                            '' citation: "http://forums.msdn.microsoft.com/en-US/netfxbcl/thread/48b4a763-7387-46da-8fc2-3e885670f62c"
                            '' so i came up with my own work around.
                            '' They spoke of using delegates, but i didn't need it that extensive for what I am trying to do.
                            '' 
                            PENDING_REQUEST_INFO.Who = who
                            PENDING_REQUEST_INFO.Auto_RESPONSE = response
                            PENDING_REQUEST_INFO.LineHoldingOn = line
                            PENDING_REQUEST_INFO.XFER_NOTES = notes
                            PENDING_REQUEST_INFO.LeadNumber = leadnum
                            STATIC_VARIABLES.PendingXFER = True
                            'Transfer_Lead2.Show()
                            mobjText = New StringBuilder()
                            Exit Select
                        Case Is = "RESPONSE"
                            Select Case MSG
                                Case Is = "ACCEPT"
                                    '' from here run a query to find the default startup form to populate the information 
                                    '' that was passed along. 
                                    'MsgBox("ACCEPT of xfer.")
                                    PENDING_REQUEST_INFO.GetStartupForm = True
                                    mobjText = New StringBuilder()
                                    Exit Select
                                Case Is = "DECLINE"
                                    MsgBox("User has Declined your xfer request.", MsgBoxStyle.Information, "XFER Declined")
                                    mobjText = New StringBuilder()
                                    Exit Select
                                Case Is = "TIMEOUT"
                                    MsgBox("The pending XFER has timed out.", MsgBoxStyle.Information, "XFER Timed Out")
                                    mobjText = New StringBuilder()
                                    Exit Select
                            End Select
                            Exit Select
                        Case Is = "CHAT"
                            '' pipe information to user chat form here.
                            ''
                            '' server side code:
                            '' -----------------
                            ''Dim UserRequesting As String = GetUserName(Trim(sender.ClientGUID.ToString))
                            ''UpdateStatus(UserRequesting & " sent:\\> " & cliMSG)
                            ' ''sender.Send("| | " & cliMSG & "|" & "CHAT" & "|SRV| |")
                            ''Dim d As DictionaryEntry
                            ''Dim objclient As ClientOBJ
                            ''For Each d In ClientHashes
                            ''    objclient = d.Value
                            ''    objclient.Send("| |" & cliMSG & "|" & "CHAT" & "|SRV|" & UserRequesting.ToString & "|")
                            ''Next
                            ''Exit Select

                            '' simple stuff here. just break apart the message wich is done above, translate user name
                            '' and feed to list view.
                            '' 

                            Control.CheckForIllegalCrossThreadCalls = False
                            Dim str4 = Split(MSG.ToString, Chr(13), -1)
                            PENDING_REQUEST_INFO.Sender = Recip
                            PENDING_REQUEST_INFO.MSG = str4(0).ToString
                            PENDING_REQUEST_INFO.IncommingChatMsg = True
                            mobjText = New StringBuilder()
                            Exit Select

                    End Select
                    Exit Select
                Case Else
            End Select
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("USER_LOGIC", "ByVal Bytes() As Byte, ByVal offset As Integer, ByVal count As Integer", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Network", "BuildString")

        End Try


    End Sub
    Private Sub MarkAsDisconnected(ByVal Who As String)
        MsgBox("Server Unavailable/Disconnected.")
        'txtSend.ReadOnly = True
        'btnSend.Enabled = False
    End Sub
    Private Sub DisplayText(ByVal t As String)
        'txtDisplay.AppendText(t)
        MsgBox(t.ToString)
    End Sub
    Private Sub DoRead(ByVal ar As IAsyncResult)
        Try
            Dim intCount As Integer

            Try
                intCount = mobjClient.GetStream.EndRead(ar)
                If intCount < 1 Then
                    MarkAsDisconnected("DoRead")
                    Exit Sub
                End If

                BuildString(mardata, 0, intCount)

                mobjClient.GetStream.BeginRead(mardata, 0, 1024, AddressOf DoRead, Nothing)

                '' do read will essentiall state wether or not the stream has been reset.
                '' 

                STATIC_VARIABLES.NET_CLIENT = mobjClient

            Catch e As Exception
                MarkAsDisconnected("DoRead")
            End Try
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("USER_LOGIC", "ByVal ar As IAsyncResult", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Network", "DoRead")
        End Try


    End Sub
    Public Sub LogOut_Of_System(ByVal UserName As String, ByVal UserPWD As String)
        Try
            If UserName.ToString.Length < 2 Then
                Application.Exit()
                End
            End If
            If UserName.ToString = "Admin" Then
                UserName = "Admin"
            End If
            'Dim str = Split(UserName, " ", 2)
            'Dim lname As String = ""
            Dim fname As String = ""
            fname = UserName
            'lname = str(1)

            Dim cmdUP As SqlCommand = New SqlCommand("UPDATE iss.dbo.userpermissiontable " _
            & " SET LoggedOn = 0, " _
            & "     MachineName = ' ', " _
            & "     IP = ' ' " _
            & " WHERE UserName = @USR and UserPWD = @PWD", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@USR", fname)
            ' Dim param33 As SqlParameter = New SqlParameter("@USRL", lname)
            Dim param2 As SqlParameter = New SqlParameter("@PWD", UserPWD)
            cmdUP.Parameters.Add(param1)
            cmdUP.Parameters.Add(param2)
            ' cmdUP.Parameters.Add(param33)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdUP.ExecuteReader(CommandBehavior.CloseConnection)
            r1.Close()
            cnn.Close()
            Main.TSMain.Enabled = False
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("USER_LOGIC", "ByVal UserName As String, ByVal UserPWD As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "LogOut_Of_System")
        End Try
    End Sub
    Public Sub CheckPermissions()

        If STATIC_VARIABLES.Administration = True Then
            Main.tsbadmin.Enabled = True
        ElseIf STATIC_VARIABLES.Administration = False Then
            Main.tsbadmin.Enabled = False
        End If

        If STATIC_VARIABLES.ColdCall = True Then
            Main.ColdCallingToolStripMenuItem.Enabled = True
        ElseIf STATIC_VARIABLES.ColdCall = False Then
            Main.ColdCallingToolStripMenuItem.Enabled = False
        End If

        If STATIC_VARIABLES.WarmCall = True Then
            Main.WarmCallingToolStripMenuItem.Enabled = True
        ElseIf STATIC_VARIABLES.WarmCall = False Then
            Main.WarmCallingToolStripMenuItem.Enabled = False
        End If

        If STATIC_VARIABLES.Recovery = True Then
            Main.RecoveryToolStripMenuItem.Enabled = True
        ElseIf STATIC_VARIABLES.Recovery = False Then
            Main.RecoveryToolStripMenuItem.Enabled = False
        End If

        If STATIC_VARIABLES.PreviousCust = True Then
            Main.PreviousCustomersToolStripMenuItem.Enabled = True
        ElseIf STATIC_VARIABLES.PreviousCust = False Then
            Main.PreviousCustomersToolStripMenuItem.Enabled = False
        End If

        If STATIC_VARIABLES.Install = True Then
            Main.tsbinstall.Enabled = True
        ElseIf STATIC_VARIABLES.Install = False Then
            Main.tsbinstall.Enabled = False
        End If

        If STATIC_VARIABLES.Finance = True Then
            Main.tsbfinance.Enabled = True
        ElseIf STATIC_VARIABLES.Finance = False Then
            Main.tsbfinance.Enabled = False
        End If

        If STATIC_VARIABLES.MarketingManager = True Then
            Main.tsbmarketing.Enabled = True
        ElseIf STATIC_VARIABLES.MarketingManager = False Then
            Main.tsbmarketing.Enabled = False
        End If

        If STATIC_VARIABLES.SalesManager = True Then
            Main.tsbsales.Enabled = True
        ElseIf STATIC_VARIABLES.SalesManager = False Then
            Main.tsbsales.Enabled = False
        End If

        If STATIC_VARIABLES.Confirmer = True Then
            Main.ConfirmingToolStripMenuItem.Enabled = True
        ElseIf STATIC_VARIABLES.Confirmer = False Then
            Main.ConfirmingToolStripMenuItem.Enabled = False
        End If

    End Sub
End Class
