Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Drawing
Imports System.Net.Sockets
Imports System.IO.FileInfo
Imports System.IO.DirectoryInfo


Public Class printingOperations
#Region "Design Notes"

    ''
    '' a) query for printable information
    ''    i) with exclusions
    ''   ii).Without exclusions
    '' b) Method to print the leads for reps without emails
    '' c) determine who can and cannot get emails.
    '' d) Print default mass list for Marketing Manager
    '' 


    ''
    '' 1) construct a print OBJ
    '' 2) send OBJ to to method to construct a styled printable page (DHMTL)
    '' 3) Show Print Preview()
    '' 4) ????
    '' 5) Profit.
    ''

    ' local test directory:>   C:\Users\Clay\Desktop\Print Leads\
    ' server directory:>       \\server.greenworks.local\Company\ISS\Print Leads

    Public Const local_test_directory_print = "C:\Users\Clay\Desktop\Print Leads\"
    Public Const Server_directory_print = "\\server.greenworks.local\Company\ISS\Print Leads"

#End Region

#Region "Structure - Print Msg Body"
    Public Structure MessageBody
        Public RecordID As String
        Public ApptDayOfWeek As String
        Public customerLName As String
        Public customer1FName As String
        Public customer2LName As String
        Public ApptDate As String
        Public ApptTime As String '' split this up or reformat to make it look better
        Public Product1 As String
        Public Product2 As String
        Public Product3 As String
        Public SpecialInstruction As String
        Public LeadGeneratedOn As String '' need to split this up to show date and not time
        Public PLS As String
        Public SLS As String
        Public Marketer As String
        Public LastMResult As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
        Public HousePhone As String
        Public AltPhone1 As String
        Public AltPhone2 As String
    End Structure
#End Region

#Region "Construct Boiler Plate HTML Page"
    Public Sub Create_STD_WEB_PAGE_FOR_PRINT(ByVal lead As MessageBody, ByVal LeadNum As String)
        Dim msg_body As MessageBody = lead
        Dim fn_nfo As New System.IO.FileInfo(local_test_directory_print & "\" & LeadNum.ToString & "-" & Date.Now.ToString & ".htm")
        Dim fs As StreamWriter
        If fn_nfo.Exists = True Then
            System.IO.File.Delete(local_test_directory_print & "\" & LeadNum.ToString & "-" & Date.Now.ToString & ".htm")
            fs = New StreamWriter(local_test_directory_print & "\" & LeadNum.ToString & "-" & Date.Now.ToString & ".htm")
        ElseIf fn_nfo.Exists = False Then
            fs = New StreamWriter(local_test_directory_print & "\" & LeadNum.ToString & "-" & Date.Now.ToString & ".htm")
        End If

        fs.WriteLine("<!DOCTYPE HTML>")
        fs.WriteLine("<head>")
        fs.WriteLine("<title>")
        fs.WriteLine("Print Lead:> " & LeadNum.ToString & " - " & Date.Today.Now.ToString)
        fs.WriteLine("</title>")
        fs.WriteLine("<style type='text/css'>span {font-family:""Verdana"",sans-serif;font-size:1.2em;font-weight:normal;}")
        fs.WriteLine("</style>")
        fs.WriteLine("</head>")
        fs.WriteLine("<body>")

        fs.WriteLine("<span>----------</span><br />")
        fs.WriteLine("<span>" & msg_body.ApptDayOfWeek & " - " & msg_body.ApptDate & ", " & msg_body.ApptTime & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>Customer ID: " & msg_body.RecordID & "</span>")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & msg_body.customerLName & ", " & msg_body.customer1FName & " & " & msg_body.customer2LName & "</span><br />")
        Dim stAddress As String = msg_body.StAddress & vbCrLf & msg_body.City & ", " & msg_body.State & " " & msg_body.Zip
        fs.WriteLine("<span>" & stAddress & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>House Phone: " & msg_body.HousePhone & "</span><br />")
        fs.WriteLine("<span>Alt Phone 1: " & msg_body.AltPhone1 & "</span><br />")
        fs.WriteLine("<span>Alt Phone 2: " & msg_body.AltPhone2 & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & msg_body.Product1 & vbCrLf & msg_body.Product2 & vbCrLf & msg_body.Product3 & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>Special Instruction: " & msg_body.SpecialInstruction & "</span><br />")
        fs.WriteLine("<span>Lead Generated On: " & msg_body.LeadGeneratedOn & "</span><br />")
        fs.WriteLine("<span>Primary Lead Source: " & msg_body.PLS & "</span><br />")
        fs.WriteLine("<span>Secondary Lead Source: " & msg_body.SLS & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & msg_body.LastMResult & "</span><br />")
        fs.WriteLine("<span>----------</span><br />")


        fs.WriteLine("</body>")
        fs.WriteLine("</html>")



    End Sub
#End Region
End Class
