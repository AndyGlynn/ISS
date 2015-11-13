Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Sockets


Public Class USER_LOGICv2

    ''
    '' creation data: 8-15-2015 @ 928am
    '' author: a.clay
    ''
    '' login to system
    '' log out of system
    '' get ipv4(s)
    '' get startup form
    '' get startup form dimensions 
    '' -- no need for maximaxized flag due to dimensions and resizing ie if screen dimension = 1024x768 screen is maximized on 1024 monitor
    '' -- can add for redundancy 
    '' get permission(s)
    '' get machine name
    ''

    ''
    '' structs: employee,LAST_FORM_AND_SCREEN_SIZE
    ''
    Public Structure Employee
        Public ID As String
        Public FName As String
        Public LName As String
        Public MName As String
        Public ScreenPositionX As String
        Public ScreenPositionY As String
        Public ScreenWidth As String
        Public ScreenHeight As String
        Public EmployeeID As String
        Public LastForm As String
        Public LastQuery As String
        Public LastLogon As String
        Public HostMachineName As String
        Public HostMachineIPV4 As String
        Public LastPasswordUsed As String
        Public CurrentPassword As String
        Public Department As String
        Public CanEnterLead As Boolean
        Public CanViewAdmin As Boolean
        Public CanViewPrint As Boolean
        Public CanViewIssueLeads As Boolean
        Public CanPrint As Boolean
        Public CanBrowseInternet As Boolean
        Public CanViewPhoneNumber As Boolean
        Public LoggedOn As Boolean
        Public ColdCall As Boolean
        Public WarmCall As Boolean
        Public PreviousCust As Boolean
        Public RecoveryLeads As Boolean
        Public Confirmer As Boolean
        Public SalesManager As Boolean
        Public MarketingManager As Boolean
        Public Finance As Boolean
        Public Install As Boolean
        Public Administration As Boolean
        Public LIC_KEY As String
        Public LEASE_KEY As String
        Public Has_Email As Boolean
        Public Has_SMS As Boolean
        Public PhoneNumber As String
        Public EmailAddress As String
    End Structure

    Public Structure LAST_FORM_AND_SCREEN_SIZE
        Public ScreenX As String
        Public ScreenY As String
        Public ScreenH As String
        Public ScreenW As String
        Public FormName As String
    End Structure

    ''
    '' consts: sql cnx
    '' 
    Public Const cnx As String = "Server=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1"

    ''
    '' props:
    '' 

    ''
    '' methods / subs / funcs
    '' 
#Region "Get IPV4s"
    Public arIPV4s
    Public ReadOnly Property IPV4_LIST As ArrayList
        Get
            Return arIPV4s
        End Get
    End Property
    Public Function GET_arIPV4s_FROM_LOCAL_MACHINE() As ArrayList
        Dim arIPs() = Dns.GetHostAddresses(My.Computer.Name.ToString)
        arIPV4s = New ArrayList
        Dim x
        For Each x In arIPs
            If x.AddressFamily = AddressFamily.InterNetwork Then
                arIPV4s.add(x.ToString)
            End If
        Next
        Return arIPV4s
    End Function
#End Region

#Region "Get Host Machine Name"
    Public Function Get_MACH_NAME()
        Dim name As String = My.Computer.Name.ToString
        Return name
    End Function
#End Region

#Region "Login To System"
    Public Function Check_User_Exists(ByVal FName As String, ByVal LName As String)
        Dim Exists As Boolean = False
        Dim usr_cnx As SqlConnection = New SqlConnection(cnx)
        Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from tblTestEmployee where FName = '" & FName & "' and LName = '" & LName & "';", usr_cnx)
        usr_cnx.Open()
        Dim cnt As Integer = cmdCNT.ExecuteScalar
        usr_cnx.Close()
        usr_cnx = Nothing
        If cnt = 0 Then
            Exists = False
        ElseIf cnt = 1 Then
            Exists = True
        End If
        Return Exists
    End Function

    Public Function Check_Password(ByVal PWD As String, ByVal FName As String, ByVal LName As String)
        Dim connect_or_not As Boolean = False

        Dim num_errors As Integer = 0

        If PWD.ToString.Length <= 0 Then
            num_errors += 1
        End If

        If FName.ToString.Length <= 0 Then
            num_errors += 1
        End If

        If LName.ToString.Length <= 0 Then
            num_errors += 1
        End If

        If num_errors >= 1 Then
            MsgBox("Please check your entry and try again. " & vbCrLf & "You need to make sure you First Name, Last Name, and Password are not blank.", MsgBoxStyle.Critical, "Check Errors")
            Return connect_or_not
            Exit Function
        ElseIf num_errors <= 0 Then
            
            Dim usr_cnx As SqlConnection = New SqlConnection(cnx)
            Dim cmdCheckPW As SqlCommand = New SqlCommand("SELECT count(ID) from tblTestEmployee where FName = '" & FName & "' and LName = '" & LName & "' and CurrentPassword = '" & PWD & "';", usr_cnx)
            usr_cnx.Open()
            Dim validation_cnt As Integer = 0
            validation_cnt = cmdCheckPW.ExecuteScalar
            usr_cnx.Close()
            usr_cnx = Nothing
            If validation_cnt = 1 Then
                connect_or_not = True
            ElseIf validation_cnt <= 0 Then
                connect_or_not = False
            End If
            Return connect_or_not
        End If

        

    End Function
#Region "Update Last Login Info - OVERLOADED"
    Public Sub Update_LAST_LOGIN(ByVal FName As String, ByVal LName As String, ByVal PWD As String)
        Dim usr_cnx As SqlConnection = New SqlConnection(cnx)
        usr_cnx.Open()
        Dim LastLogin As String = Date.Now.ToString
        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE tblTestEmployee set LastLogin = '" & LastLogin & "' WHERE FName = '" & FName & "' and LName = '" & LName & "';", usr_cnx)
        cmdUP.ExecuteScalar()
        usr_cnx.Close()
        usr_cnx = Nothing

    End Sub
    Public Sub Update_LAST_LOGIN(ByVal EmployeeID As String, ByVal PWD As String)
        Dim usr_cnx As SqlConnection = New SqlConnection(cnx)
        usr_cnx.Open()
        Dim LastLogin As String = Date.Now.ToString
        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE tblTestEmployee set LastLogin = '" & LastLogin & "' WHERE EmployeeID = '" & EmployeeID & "';", usr_cnx)
        cmdUP.ExecuteScalar()
        usr_cnx.Close()
        usr_cnx = Nothing
    End Sub
#End Region
    Function Get_User_Obj(ByVal FName As String, ByVal LName As String, ByVal PWD As String)
        Dim usr_cnx As SqlConnection = New SqlConnection(cnx)
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT * from tblTestEmployee where FName = '" & FName & "' and LName = '" & LName & "' and CurrentPassword= '" & PWD & "'", usr_cnx)
        usr_cnx.Open()
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        Dim g As New Employee
        While r1.Read
            g.ID = r1.Item("ID")
            g.FName = r1.Item("FName")
            g.LName = r1.Item("LName")
            g.MName = r1.Item("MName")
            g.ScreenPositionX = r1.Item("ScreenPositionX")
            g.ScreenPositionY = r1.Item("ScreenPositionY")
            g.ScreenWidth = r1.Item("ScreenWidth")
            g.ScreenHeight = r1.Item("ScreenHeight")
            g.EmployeeID = r1.Item("EmployeeID")
            g.LastForm = r1.Item("LastForm")
            g.LastQuery = r1.Item("LastQuery")
            g.LastLogon = r1.Item("LastLogin")
            g.HostMachineName = My.Computer.Name.ToString
            g.HostMachineIPV4 = r1.Item("HostMachineIPV4")
            g.LastPasswordUsed = r1.Item("LastPasswordUsed")
            g.CurrentPassword = r1.Item("CurrentPassword")
            g.Department = r1.Item("Department")
            g.CanEnterLead = r1.Item("CanEnterLead")
            g.CanViewAdmin = r1.Item("CanViewAdmin")
            g.CanViewPrint = r1.Item("CanViewPrint")
            g.CanViewIssueLeads = r1.Item("CanViewIssueLeads")
            g.CanPrint = r1.Item("CanPrint")
            g.CanBrowseInternet = r1.Item("CanBrowseInternet")
            g.CanViewPhoneNumber = r1.Item("CanViewPhoneNumber")
            g.LoggedOn = r1.Item("LoggedOn")
            g.ColdCall = r1.Item("ColdCall")
            g.WarmCall = r1.Item("WarmCall")
            g.PreviousCust = r1.Item("PreviousCust")
            g.RecoveryLeads = r1.Item("RecoveryLeads")
            g.Confirmer = r1.Item("Confirmer")
            g.SalesManager = r1.Item("SalesManager")
            g.MarketingManager = r1.Item("MarketingManager")
            g.Finance = r1.Item("Finance")
            g.Install = r1.Item("Install")
            g.Administration = r1.Item("Administration")
            g.LIC_KEY = r1.Item("LIC_KEY")
            g.LEASE_KEY = r1.Item("LEASE_KEY")
            g.Has_Email = r1.Item("HasEmail")
            g.Has_SMS = r1.Item("HasSMS")
            g.PhoneNumber = r1.Item("PhoneNumber")
            g.EmailAddress = r1.Item("EmailAddress")
        End While
        r1.Close()
        usr_cnx.Close()
        usr_cnx = Nothing
        Return g
    End Function

#End Region


#Region "Log Out Of System - OVERLOADED"
    Public Sub LogOut(ByVal FName As String, ByVal LName As String, ByVal ScreenX As String, ByVal ScreenY As String, ByVal ScreenW As String, ByVal ScreenH As String, ByVal LastForm As String, ByVal LastQuery As String, _
                      ByVal MachineName As String, ByVal PWD As String)
        Dim usr_CNX As SqlConnection = New SqlConnection(cnx)
        usr_CNX.Open()
        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE tblTestEmployee set ScreenPositionX = '" & ScreenX & "', ScreenPositionY = '" & ScreenY & "', screenwidth = '" & ScreenW & "', screenheight = '" & ScreenH & "', LastForm = '" & LastForm & "', LastQuery = '" & LastQuery & "', HostMachineName = '" & Environment.MachineName.ToString & "', LastPasswordUsed = '" & PWD & "' WHERE FName = '" & FName & "' and LName ='" & LName & "';", usr_CNX)
        cmdUP.ExecuteScalar()
        usr_CNX.Close()
        usr_CNX = Nothing
        Update_LAST_LOGIN(FName, LName, PWD)
    End Sub
    Public Sub LogOut(ByVal EmployeeID As String, ByVal ScreenX As String, ByVal ScreenY As String, ByVal ScreenW As String, ByVal ScreenH As String, ByVal LastForm As String, ByVal LastQuery As String, _
                      ByVal MachineName As String, ByVal PWD As String)
        Dim usr_CNX As SqlConnection = New SqlConnection(cnx)
        usr_CNX.Open()
        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE tblTestEmployee set ScreenPositionX = '" & ScreenX & "', ScreenPositionY = '" & ScreenY & "', screenwidth = '" & ScreenW & "', screenheight = '" & ScreenH & "', LastForm = '" & LastForm & "', LastQuery = '" & LastQuery & "', HostMachineName = '" & Environment.MachineName.ToString & "', LastPasswordUsed = '" & PWD & "' WHERE EmployeeID = '" & EmployeeID & "';", usr_CNX)
        cmdUP.ExecuteScalar()
        usr_CNX.Close()
        usr_CNX = Nothing
        Update_LAST_LOGIN(EmployeeID, PWD)
    End Sub
#End Region
#Region "Get Last Form and SIZE - OVERLOADED"
    Public Function Get_LAST_FORM_AND_SIZES(ByVal Fname As String, ByVal LName As String, ByVal PWD As String)
        Dim last_form_name As String
        Dim scrX As String
        Dim scrY As String
        Dim scrH As String
        Dim scrW As String
        Dim usr_cnx As SqlConnection = New SqlConnection(cnx)
        usr_cnx.Open()
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT LastForm,ScreenPositionX,ScreenPositionY,ScreenHeight,ScreenWidth from tblTestEmployee WHERE FName = '" & Fname & "' and LName = '" & LName & "';", usr_cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        Dim z As New LAST_FORM_AND_SCREEN_SIZE
        While r1.Read
            z.FormName = r1.Item("LastForm")
            z.ScreenX = r1.Item("ScreenPositionX")
            z.ScreenY = r1.Item("ScreenPositionY")
            z.ScreenH = r1.Item("ScreenHeight")
            z.ScreenW = r1.Item("ScreenWidth")
        End While
        r1.Close()
        usr_cnx.Close()
        usr_cnx = Nothing
        Return z
    End Function
    Public Function Get_LAST_FORM_AND_SIZES(ByVal EmployeeID As String, ByVal PWD As String)
        Dim last_form_name As String
        Dim scrX As String
        Dim scrY As String
        Dim scrH As String
        Dim scrW As String
        Dim usr_cnx As SqlConnection = New SqlConnection(cnx)
        usr_cnx.Open()
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT LastForm,ScreenPositionX,ScreenPositionY,ScreenHeight,ScreenWidth from tblTestEmployee WHERE EmployeeID = '" & EmployeeID & "';", usr_cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        Dim z As New LAST_FORM_AND_SCREEN_SIZE
        While r1.Read
            z.FormName = r1.Item("LastForm")
            z.ScreenX = r1.Item("ScreenPositionX")
            z.ScreenY = r1.Item("ScreenPositionY")
            z.ScreenH = r1.Item("ScreenHeight")
            z.ScreenW = r1.Item("ScreenWidth")
        End While
        r1.Close()
        usr_cnx.Close()
        usr_cnx = Nothing
        Return z
    End Function
#End Region



End Class
