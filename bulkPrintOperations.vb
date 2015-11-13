Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.String
Imports System.Text.StringBuilder


Public Class bulkPrintOperations
    Public Const sqlCNX As String = "SERVER=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1"
    Public Const local_test_directory_print = "C:/Users/Clay/Desktop/Print Leads/"
    Public Const Server_directory_print = "\\192.168.1.2\Company\ISS\Print Leads\"


    Public ListOfLeadNumbers As ArrayList
    Public Exclu_ As Exclusions
    Public Structure Exclusions
        Public Generated As Boolean
        Public Marketer As Boolean
        Public PLS As Boolean
        Public SLS As Boolean
        Public LastMResult As Boolean
        Public Phone As Boolean
    End Structure

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


    Public Function GetLeadNumbers()
        ListOfLeadNumbers = New ArrayList
        Dim cnx_ld As New SqlConnection(sqlCNX)
        Dim cmdGet As SqlCommand = New SqlCommand("SELECT distinct(ID) from EnterLead;", cnx_ld)
        cnx_ld.Open()
        Dim r1 As SqlDataReader = cmdGet.ExecuteReader
        While r1.Read
            ListOfLeadNumbers.Add(r1.Item("ID"))
        End While
        r1.Close()
        cnx_ld.Close()
        cnx_ld = Nothing
        Return ListOfLeadNumbers
    End Function
    Public Sub PopulateListViewWithLeadNumbers()
        Dim lst_ As New ArrayList
        lst_ = GetLeadNumbers()
        Dim ls_col As ListView.ListViewItemCollection = New ListView.ListViewItemCollection(Form1.ListView1)
        Dim g As Integer = 0
        For g = 0 To lst_.Count - 1
            Dim lvItem As New ListViewItem
            lvItem.Text = lst_(g).ToString
            ls_col.Add(lvItem)
        Next

    End Sub
    Public Sub New()

    End Sub

#Region "Get Exclusions"
    Public Function GetExclusions()
        Dim cnx_ex As SqlConnection = New SqlConnection(sqlCNX)
        Dim cmdGet As SqlCommand = New SqlCommand("SELECT * From iss.dbo.Exclusions", cnx_ex)
        cnx_ex.Open()
        Dim exclusion_ As New Exclusions
        Dim r1 As SqlDataReader = cmdGet.ExecuteReader
        While r1.Read
            exclusion_.Generated = r1.Item("Generated")
            exclusion_.Marketer = r1.Item("Marketer")
            exclusion_.PLS = r1.Item("PLS")
            exclusion_.SLS = r1.Item("SLS")
            exclusion_.LastMResult = r1.Item("LastMResult")
            exclusion_.Phone = r1.Item("Phone")
        End While
        r1.Close()
        cnx_ex.Close()
        cnx_ex = Nothing
        Return exclusion_
    End Function
    Public Property ExclusionList As Exclusions
        Get
            Return Exclu_
        End Get
        Set(value As Exclusions)
            Exclu_ = value
        End Set
    End Property
#End Region

    Public Sub DoTheWork(ByVal LeadNum As String)
        Dim msg_body As MessageBody = GenerateMSG_BODY(LeadNum)
        GenerateBoiler(msg_body, LeadNum)
        frmPrint.wbPrint.Navigate(Server_directory_print & LeadNum.ToString & ".htm")

    End Sub
    Public Sub DoTheWork_EXCLUSIONS(ByVal LeadNum As String, ByVal Exclusions As Exclusions)
        Dim msg_body As MessageBody = GenerateMSG_BODY(LeadNum)
        GenerateBoiler_Exclusions(msg_body, LeadNum, Exclusions)
        frmPrint.wbPrint.Navigate(Server_directory_print & LeadNum.ToString & ".htm")
    End Sub





    Public Function GenerateMSG_BODY(ByVal LeadNum As String)
        Dim msg_ As New MessageBody

        Dim cnx_MSG As SqlConnection = New SqlConnection(sqlCNX)
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT ID,ApptDay,Contact1FirstName,Contact2FirstName,Contact1LastName,ApptDate,ApptTime,Product1,Product2,Product3," _
        & "SpecialInstruction,LeadGeneratedOn,PrimaryLeadSource,SecondaryLeadSource,Marketer,MarketingResults,StAddress,City,State,Zip," _
        & "HousePhone,AltPhone1,AltPhone2 from EnterLead WHERE ID = '" & LeadNum & "';", cnx_MSG)
        cnx_MSG.Open()
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            msg_.RecordID = r1.Item("ID")
            msg_.ApptDayOfWeek = r1.Item("ApptDay")
            msg_.customer1FName = r1.Item("Contact1FirstName")
            msg_.customer2LName = r1.Item("Contact2FirstName")
            msg_.customerLName = r1.Item("Contact1LastName")
            msg_.ApptDate = r1.Item("ApptDate")
            msg_.ApptTime = r1.Item("ApptTime")
            msg_.Product1 = r1.Item("Product1")
            msg_.Product2 = r1.Item("Product2")
            msg_.Product3 = r1.Item("Product3")
            msg_.SpecialInstruction = r1.Item("SpecialInstruction")
            msg_.LeadGeneratedOn = r1.Item("LeadGeneratedOn")
            msg_.PLS = r1.Item("PrimaryLeadSource")
            msg_.SLS = r1.Item("SecondaryLeadSource")
            msg_.Marketer = r1.Item("Marketer")
            msg_.LastMResult = GetLastKnownMarketingResult(LeadNum)
            msg_.StAddress = r1.Item("StAddress")
            msg_.City = r1.Item("City")
            msg_.State = r1.Item("State")
            msg_.Zip = r1.Item("Zip")
            msg_.HousePhone = r1.Item("HousePhone")
            msg_.AltPhone1 = r1.Item("AltPhone1")
            msg_.AltPhone2 = r1.Item("AltPhone2")
        End While
        r1.Close()
        cnx_MSG.Close()
        cnx_MSG = Nothing
        Return msg_
    End Function

    Private Function GetLastKnownMarketingResult(ByVal LeadNum As String)
        Dim cmdMRes As SqlConnection = New SqlConnection(sqlCNX)
        cmdMRes.Open()
        Dim cmdDO As SqlCommand = New SqlCommand("LastKnownMarketingResult", cmdMRes)
        cmdDO.CommandType = CommandType.StoredProcedure
        cmdDO.Parameters.AddWithValue("@ID", LeadNum)
        Dim res As String = cmdDO.ExecuteScalar
        cmdMRes.Close()
        cmdMRes = Nothing
        Return res
    End Function

    Public Sub GenerateBoiler_Exclusions(ByVal MSG_BODY As MessageBody, ByVal LeadNum As String, ByVal Exclusions As Exclusions)
        Dim _msg As MessageBody = MSG_BODY
        Dim dt As String = Date.Today.ToString
        Dim ex_set As Exclusions = Exclusions
        Dim fn_nfo As New System.IO.FileInfo(Server_directory_print & "\" & LeadNum.ToString & ".htm")
        Dim fs As StreamWriter
        If fn_nfo.Exists = True Then
            System.IO.File.Delete(Server_directory_print & "\" & LeadNum.ToString & ".htm")
            fs = New StreamWriter(Server_directory_print & "\" & LeadNum.ToString & ".htm")
        ElseIf fn_nfo.Exists = False Then
            fs = New StreamWriter(Server_directory_print & "\" & LeadNum.ToString & ".htm")
        End If

        fs.WriteLine("<!DOCTYPE HTML>")
        fs.WriteLine("<head>")
        fs.WriteLine("<title>")
        fs.WriteLine("Print Lead:> " & LeadNum.ToString & " - " & Date.Now.ToString)
        fs.WriteLine("</title>")
        fs.WriteLine("<style type='text/css'>span {font-family:""Verdana"",sans-serif;font-size:1.2em;font-weight:normal;}")
        fs.WriteLine("</style>")
        fs.WriteLine("</head>")
        fs.WriteLine("<body>")

        fs.WriteLine("<span>----------</span><br />")

        Dim time As String = SplitTime(MSG_BODY.ApptTime)

        fs.WriteLine("<span>" & MSG_BODY.ApptDayOfWeek & " - " & MSG_BODY.ApptDate & ", " & time & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>Customer ID: " & MSG_BODY.RecordID & "</span>")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & MSG_BODY.customerLName & ", " & MSG_BODY.customer1FName & " & " & MSG_BODY.customer2LName & "</span><br />")
        Dim stAddress As String = MSG_BODY.StAddress & vbCrLf & MSG_BODY.City & ", " & MSG_BODY.State & " " & MSG_BODY.Zip
        fs.WriteLine("<span>" & stAddress & "</span><br />")
        fs.WriteLine("<br />")


        If ex_set.Phone = True Then
            ''' do nothing
        ElseIf ex_set.Phone = False Then
            fs.WriteLine("<span>House Phone: " & MSG_BODY.HousePhone & "</span><br />")
            fs.WriteLine("<span>Alt Phone 1: " & MSG_BODY.AltPhone1 & "</span><br />")
            fs.WriteLine("<span>Alt Phone 2: " & MSG_BODY.AltPhone2 & "</span><br />")
        End If

        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & MSG_BODY.Product1 & vbCrLf & MSG_BODY.Product2 & vbCrLf & MSG_BODY.Product3 & "</span><br />")
        fs.WriteLine("<br />")

        fs.WriteLine("<span>Special Instruction: " & MSG_BODY.SpecialInstruction & "</span><br />")

        If ex_set.Generated = True Then
            '' do nothing 
        ElseIf ex_set.Generated = False Then
            fs.WriteLine("<span>Lead Generated On: " & MSG_BODY.LeadGeneratedOn & "</span><br />")
        End If

        If ex_set.PLS = True Then
            '' do nothing 
        ElseIf ex_set.PLS = False Then
            fs.WriteLine("<span>Primary Lead Source: " & MSG_BODY.PLS & "</span><br />")
        End If

        If ex_set.SLS = True Then
            '' do nothing 
        ElseIf ex_set.SLS = False Then
            fs.WriteLine("<span>Secondary Lead Source: " & MSG_BODY.SLS & "</span><br />")
        End If

        If ex_set.Marketer = True Then
            '' do nothing
        ElseIf ex_set.Marketer = False Then
            fs.WriteLine("<span>Marketer : " & MSG_BODY.Marketer & "</span><br />")
        End If

        fs.WriteLine("<br />")
        If ex_set.LastMResult = True Then
        ElseIf ex_set.LastMResult = False Then
            fs.WriteLine("<span>" & MSG_BODY.LastMResult & "</span><br />")
        End If

        fs.WriteLine("<span>----------</span><br />")


        fs.WriteLine("</body>")
        fs.WriteLine("</html>")
        fs.Flush()
        fs.Close()

    End Sub


    Public Sub GenerateBoiler(ByVal MSG_BODY As MessageBody, ByVal LeadNum As String)
        Dim _msg As MessageBody = MSG_BODY
        Dim dt As String = Date.Today.ToString

        Dim fn_nfo As New System.IO.FileInfo(Server_directory_print & "\" & LeadNum.ToString & ".htm")
        Dim fs As StreamWriter
        If fn_nfo.Exists = True Then
            System.IO.File.Delete(Server_directory_print & "\" & LeadNum.ToString & ".htm")
            fs = New StreamWriter(Server_directory_print & "\" & LeadNum.ToString & ".htm")
        ElseIf fn_nfo.Exists = False Then
            fs = New StreamWriter(Server_directory_print & "\" & LeadNum.ToString & ".htm")
        End If

        fs.WriteLine("<!DOCTYPE HTML>")
        fs.WriteLine("<head>")
        fs.WriteLine("<title>")
        fs.WriteLine("Print Lead:> " & LeadNum.ToString & " - " & Date.Now.ToString)
        fs.WriteLine("</title>")
        fs.WriteLine("<style type='text/css'>span {font-family:""Verdana"",sans-serif;font-size:1.2em;font-weight:normal;}")
        fs.WriteLine("</style>")
        fs.WriteLine("</head>")
        fs.WriteLine("<body>")

        fs.WriteLine("<span>----------</span><br />")

        Dim time As String = SplitTime(MSG_BODY.ApptTime)

        fs.WriteLine("<span>" & MSG_BODY.ApptDayOfWeek & " - " & MSG_BODY.ApptDate & ", " & time & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>Customer ID: " & MSG_BODY.RecordID & "</span>")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & MSG_BODY.customerLName & ", " & MSG_BODY.customer1FName & " & " & MSG_BODY.customer2LName & "</span><br />")
        Dim stAddress As String = MSG_BODY.StAddress & vbCrLf & MSG_BODY.City & ", " & MSG_BODY.State & " " & MSG_BODY.Zip
        fs.WriteLine("<span>" & stAddress & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>House Phone: " & MSG_BODY.HousePhone & "</span><br />")
        fs.WriteLine("<span>Alt Phone 1: " & MSG_BODY.AltPhone1 & "</span><br />")
        fs.WriteLine("<span>Alt Phone 2: " & MSG_BODY.AltPhone2 & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & MSG_BODY.Product1 & vbCrLf & MSG_BODY.Product2 & vbCrLf & MSG_BODY.Product3 & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>Special Instruction: " & MSG_BODY.SpecialInstruction & "</span><br />")
        fs.WriteLine("<span>Lead Generated On: " & MSG_BODY.LeadGeneratedOn & "</span><br />")
        fs.WriteLine("<span>Primary Lead Source: " & MSG_BODY.PLS & "</span><br />")
        fs.WriteLine("<span>Secondary Lead Source: " & MSG_BODY.SLS & "</span><br />")
        fs.WriteLine("<span>Marketer: " & MSG_BODY.Marketer & "</span><br />")
        fs.WriteLine("<br />")
        fs.WriteLine("<span>" & MSG_BODY.LastMResult & "</span><br />")
        fs.WriteLine("<span>----------</span><br />")


        fs.WriteLine("</body>")
        fs.WriteLine("</html>")
        fs.Flush()
        fs.Close()

    End Sub

    Public Function SplitTime(ByVal ApptTime As String)
        Dim time As String = ApptTime
        Dim peices() = time.Split(" ")
        Dim secondary = peices(1).split(":")
        Dim timeConstructed As String = secondary(0) & ":" & secondary(1) & " " & peices(2)
        Return timeConstructed
    End Function

    Public Function Generate_BULK_MSG_BODY(ByVal LeadNums As ArrayList)
        Dim strMSG As String = "<!DOCTYPE html><head><title>Bulk Print For" & Date.Now.ToString & "</title><style type='text/css'>span{font-family:'Verdana',sans-serif;font-size:1.2em;color:black;font-weight:normal;}</style></head><span>----BEGIN LEADS SELECTED----</span><br />" & vbCrLf
        Dim arMSGs As New ArrayList '' holds a 'list ' of selected msgs in messagebody structs

        Dim y As Integer = 0
        For y = 0 To LeadNums.Count - 1
            arMSGs.Add(GenerateMSG_BODY(LeadNums(y)))
        Next

        ''
        '' now that we have 'list', convert a lines of each msg to string and drop to string variable
        '' to send the ONE variable for print object / job
        ''

        Dim a As Integer = 0
        For a = 0 To arMSGs.Count - 1
            Dim c As MessageBody = arMSGs(a)
            strMSG += "<span>------------</span><br />" & vbCrLf
            strMSG += "<span>" & c.ApptDayOfWeek & " - " & c.ApptDate & ", " & SplitTime(c.ApptTime) & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf
            strMSG += "<span>Customer ID: " & c.RecordID & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf
            strMSG += "<span>" & c.customerLName & ", " & c.customer1FName & " & " & c.customer2LName & "</span><br />" & vbCrLf
            Dim stAddress As String = c.StAddress & vbCrLf & c.City & ", " & c.State & " " & c.Zip
            strMSG += "<span>" & stAddress & "</span><br />" & vbCrLf
            strMSG += "<br />"
            strMSG += "<span>House Phone: " & c.HousePhone & "</span><br />" & vbCrLf
            strMSG += "<span>Alt Phone 1: " & c.AltPhone1 & "</span><br />" & vbCrLf
            strMSG += "<span>Alt Phone 2: " & c.AltPhone2 & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf
            strMSG += "<span>" & c.Product1 & vbCrLf & c.Product2 & vbCrLf & c.Product3 & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf
            strMSG += "<span>Lead Generated On: " & c.LeadGeneratedOn & "</span><br />" & vbCrLf
            strMSG += "<span>Special Instruction: " & c.SpecialInstruction & "</span><br />" & vbCrLf
            strMSG += "<span>Primary Lead Source: " & c.PLS & "</span><br />" & vbCrLf
            strMSG += "<span>Secondary Lead Source: " & c.SLS & "</span><br />" & vbCrLf
            strMSG += "<span>Marketer: " & c.Marketer & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf
            strMSG += "<span>" & GetLastKnownMarketingResult(c.RecordID) & "</span><br />" & vbCrLf
            strMSG += "<span>------------</span>" & vbCrLf
        Next
        strMSG += "<span>----END LEADS SELECTED----</span>" & vbCrLf
        strMSG += "</body>" & vbCrLf
        strMSG += "</html>" & vbCrLf

        Return strMSG

    End Function

    Public Function Generate_BULK_MSG_BODY_EXCLUSIONS(ByVal LeadNums As ArrayList, ByVal Exclusions As Exclusions)
        Dim strMSG As String = "<!DOCTYPE html><head><title>Bulk Print For" & Date.Now.ToString & "</title><style type='text/css'>span{font-family:'Verdana',sans-serif;font-size:1.2em;color:black;font-weight:normal;}</style></head><span>----BEGIN LEADS SELECTED----</span><br />" & vbCrLf
        Dim arMSGs As New ArrayList '' holds a 'list ' of selected msgs in messagebody structs

        Dim y As Integer = 0
        For y = 0 To LeadNums.Count - 1
            arMSGs.Add(GenerateMSG_BODY(LeadNums(y)))
        Next

        Dim ex_set As bulkPrintOperations.Exclusions
        ex_set = GetExclusions()

        ''
        '' now that we have 'list', convert a lines of each msg to string and drop to string variable
        '' to send the ONE variable for print object / job
        ''

        Dim a As Integer = 0
        For a = 0 To arMSGs.Count - 1
            Dim c As MessageBody = arMSGs(a)
            strMSG += "<span>------------</span><br />" & vbCrLf
            strMSG += "<span>" & c.ApptDayOfWeek & " - " & c.ApptDate & ", " & SplitTime(c.ApptTime) & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf
            strMSG += "<span>Customer ID: " & c.RecordID & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf
            strMSG += "<span>" & c.customerLName & ", " & c.customer1FName & " & " & c.customer2LName & "</span><br />" & vbCrLf
            Dim stAddress As String = c.StAddress & vbCrLf & c.City & ", " & c.State & " " & c.Zip
            strMSG += "<span>" & stAddress & "</span><br />" & vbCrLf
            strMSG += "<br />"

            If ex_set.Phone = True Then
                '' do nothing 
            ElseIf ex_set.Phone = False Then
                strMSG += "<span>House Phone: " & c.HousePhone & "</span><br />" & vbCrLf
                strMSG += "<span>Alt Phone 1: " & c.AltPhone1 & "</span><br />" & vbCrLf
                strMSG += "<span>Alt Phone 2: " & c.AltPhone2 & "</span><br />" & vbCrLf
            End If


            strMSG += "<br />" & vbCrLf
            strMSG += "<span>" & c.Product1 & vbCrLf & c.Product2 & vbCrLf & c.Product3 & "</span><br />" & vbCrLf
            strMSG += "<br />" & vbCrLf

            If ex_set.Generated = True Then
            ElseIf ex_set.Generated = False Then
                strMSG += "<span>Lead Generated On: " & c.LeadGeneratedOn & "</span><br />" & vbCrLf
            End If

            strMSG += "<span>Special Instruction: " & c.SpecialInstruction & "</span><br />" & vbCrLf
            If ex_set.PLS = True Then
            ElseIf ex_set.SLS = False Then
                strMSG += "<span>Primary Lead Source: " & c.PLS & "</span><br />" & vbCrLf
            End If

            If ex_set.SLS = True Then
            ElseIf ex_set.SLS = False Then
                strMSG += "<span>Secondary Lead Source: " & c.SLS & "</span><br />" & vbCrLf
            End If

            If ex_set.Marketer = True Then
            ElseIf ex_set.Marketer = False Then
                strMSG += "<span>Marketer: " & c.Marketer & "</span><br />" & vbCrLf
            End If

            strMSG += "<br />" & vbCrLf
            If ex_set.LastMResult = True Then
            ElseIf ex_set.LastMResult = False Then
                strMSG += "<span>" & GetLastKnownMarketingResult(c.RecordID) & "</span><br />" & vbCrLf
            End If
            strMSG += "<span>------------</span>" & vbCrLf
        Next
        strMSG += "<span>----END LEADS SELECTED----</span>" & vbCrLf
        strMSG += "</body>" & vbCrLf
        strMSG += "</html>" & vbCrLf

        Return strMSG

    End Function

    Public Sub GenerateBULK_PRINT(ByVal StrMSG As String)
        Dim dt As String = Date.Today.ToString

        Dim fn_nfo As New System.IO.FileInfo(Server_directory_print & "\BULK_PRINT.htm")
        Dim fs As StreamWriter
        If fn_nfo.Exists = True Then
            System.IO.File.Delete(Server_directory_print & "\BULK_PRINT.htm")
            fs = New StreamWriter(Server_directory_print & "\BULK_PRINT.htm")
        ElseIf fn_nfo.Exists = False Then
            fs = New StreamWriter(Server_directory_print & "\BULK_PRINT.htm")
        End If
        fs.Write(StrMSG)
        fs.Flush()
        fs.Close()
        frmPrint.wbPrint.Navigate(Server_directory_print & "BULK_PRINT.htm")
    End Sub

#Region "Can the rep get email - OVERLOADED"
    ''
    '' overload for employeeID
    ''
    Public Function CanRepGetEmail(ByVal RepFName As String, ByVal RepLName As String)
        Dim can_get_email As Boolean = False
        Dim rep_cnx As SqlConnection = New SqlConnection(sqlCNX)
        rep_cnx.Open()
        Dim cmdCheck As SqlCommand = New SqlCommand("SELECT HasEmail from tblTestEmployee where FName = '" & RepFName & "' and LName = '" & RepLName & "';", rep_cnx)
        Dim strReturn As String = cmdCheck.ExecuteScalar
        rep_cnx.Close()
        rep_cnx = Nothing
        If strReturn = "True" Then
            can_get_email = True
        ElseIf strReturn = "False" Then
            can_get_email = False
        End If
        Return can_get_email
    End Function

    Public Function CanRepGetEmail(ByVal EmployeeID As String)
        Dim can_get_email As Boolean = False
        Dim rep_cnx As SqlConnection = New SqlConnection(sqlCNX)
        rep_cnx.Open()
        Dim cmdCheck As SqlCommand = New SqlCommand("SELECT HasEmail from tblTestEmployee where EmployeeID = '" & EmployeeID & "'", rep_cnx)
        Dim strReturn As String = cmdCheck.ExecuteScalar
        rep_cnx.Close()
        rep_cnx = Nothing
        If strReturn = "True" Then
            can_get_email = True
        ElseIf strReturn = "False" Then
            can_get_email = False
        End If
        Return can_get_email
    End Function
#End Region
End Class
