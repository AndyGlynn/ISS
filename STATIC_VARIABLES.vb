Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Net.Sockets
Module STATIC_VARIABLES
    Public Const ProgramName As String = "Improveit! 360"
    Public SMN As New Collection
    Public oApp As New MapPoint.Application
    ' Public CurrentUser As String = ""
    Public CurrentManager As String = ""
    Public Cnn As String = "Data Source=192.168.1.2;Initial Catalog=iss;User Id=sa;Password=spoken1;"
    Public CnnTimedRefresh As String = "Data Source=192.168.1.2;Initial Catalog=iss;User Id=sa;Password=spoken1;"
    Public CnnCustomerHistory As String = "Data Source=192.168.1.2;Initial Catalog=iss;User Id=sa;Password=spoken1;"
    Public AttachedFilesDirectory As String = "\\SERVER\Company\ISS\Attached Files\"
    Public SAAttachedFileDirectory As String = "\\SERVER\Company\ISS\ScheduledActionAttachments\"
    Public JobPicturesFileDirectory As String = "\\SERVER\Company\ISS\Job Pictures\"
    Public EmailDirectory As String = "\\SERVER\Company\ISS\Email Leads\" & Date.Now.ToString & "\"

#Region "User Information"
    Public Login As String = ""
    Public CurrentUser As String = ""
    Public CurrentID As String = ""
    Public UserPWD As String = ""
    Public ColdCall As Boolean
    Public WarmCall As Boolean
    Public PreviousCust As Boolean
    Public Recovery As Boolean
    Public Confirmer As Boolean
    Public SalesManager As Boolean
    Public MarketingManager As Boolean
    Public Finance As Boolean
    Public Install As Boolean
    Public Administration As Boolean
    Public StartUpForm As String = ""
    Public CurrentForm As String = ""
    Public LoggedOn As Boolean
    Public DoNotShowMapping As Boolean
    Public ManagerFirstName As String = ""
    Public ManagerLastName As String = ""
    Public MachineName As String = ""
    Public IP As String = ""
    Public LicenseKey As String = "AAAAA-BBBBB-CCCCC-DDDDD-FFFFF-1"
    Public LeaseKey As String = "000-00000-001"
    Public Server_Assigned_Hash As String = ""
    Public NET_CLIENT As Tcpclient
    Public ActiveChild As Form
    Public PendingXFER As Boolean = False
    Public salesworkaround As Boolean = True

    Public employee As New List(Of USER_LOGICv2.Employee)
    Public CurrentExclusionSet As EmailIssuedLeads.Exclusions

     



#End Region
End Module

