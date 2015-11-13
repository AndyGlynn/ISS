
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Net.Sockets
Imports System.Net
Imports System.IO
Imports System.Net.Mail

Public Class EmailIssuedLeads
    Public Const cnx_string As String = "Server=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1"

    ''
    '' notes: 8-15-2015
    ''
    '' need some arrays here
    '' one array for reps that get email
    '' one array for reps that dont get email
    '' one array for leads not issued to a rep 
    '' one array for marketing manager list 
    '' 
    Public repsThatGetEmail As ArrayList
    Public repEmailAddress As ArrayList
    Public repsThatDontGetEmail As ArrayList
    Public unassignedLeads As ArrayList
    Public arMessages As ArrayList

    ''
    '' need a couple of structures
    '' 
    ''
    '' struct: employee, lead
    Public Structure RepWithEmail
        Public FName As String
        Public LName As String
        Public EmailAddress As String
    End Structure

    Public Structure Exclusions
        Public Generated As Boolean
        Public Marketer As Boolean
        Public PLS As Boolean
        Public SLS As Boolean
        Public LastMResult As Boolean
        Public Phone As Boolean
    End Structure

    Public Structure MarketingManagerOBJ
        Public FName As String
        Public LName As String
        Public EmailAddress As String
    End Structure

    Public Structure MM_ListOBJ
        Public ID As String
        Public ApptTime As String
        Public Contact1FirstName As String
        Public Contact2FirstName As String
        Public Contact1LastName As String
        Public Contact2LastName As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
        Public Product1 As String
        Public Product2 As String
        Public Product3 As String
        Public Rep1 As String
        Public Rep2 As String
        Public Phone As String
        Public AltPhone1 As String
        Public AltPhone2 As String
    End Structure

    Public Structure MessageBody
        'strLead += r1.Item("ID") & vbCrLf
        Public RecordID As String
        '          strLead += r1.Item("Marketer") & vbCrLf
        Public Marketer As String
        '       strLead += r1.Item("PrimaryLeadSource") & vbCrLf
        Public PLS As String
        '       strLead += r1.Item("SecondaryLeadSource") & vbCrLf
        Public SLS As String
        '       strLead += r1.Item("LeadGeneratedOn") & vbCrLf
        Public LeadGeneratedOn As String
        '   strLead += (r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName")) & vbCrLf
        Public C1Name As String
        '   strLead += (r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName")) & vbCrLf
        Public C2Name As String
        '   strLead += (r1.Item("StAddress") & " " & r1.Item("City") & " " & r1.Item("state") & " " & r1.Item("Zip")) & vbCrLf
        Public Address As String
        '       strLead += (r1.Item("HousePhone") & " " & r1.Item("Altphone1") & " " & r1.Item("AltPhone2")) & vbCrLf
        Public PhoneNumbers As String
        '   strLead += r1.Item("SpokeWith") & " " & r1.Item("Contact1WorkHours") & " " & r1.Item("Contact2WorkHours") & vbCrLf
        Public SpokeWithAndContactHours As String
        '   strLead += r1.Item("Product1") & " " & r1.Item("Product2") & " " & r1.Item("Product3") & r1.Item("Color") & " " & r1.Item("ProductQTY") & vbCrLf
        Public Products As String
        '   strLead += r1.Item("YearsOwned") & " " & r1.Item("HomeAge") & " " & r1.Item("HomeValue") & vbCrLf
        Public HomeInfo As String
        '   strLead += r1.Item("ApptDate") & " " & r1.Item("ApptDay") & " " & r1.Item("ApptTime") & vbCrLf & vbCrLf
        Public ApptIfno As String
        '   strLead += r1.Item("SpecialInstruction") & vbCrLf & vbCrLf
        Public SpecialInstructions As String
        '   strLead += r1.Item("Result") & vbCrLf
        Public LastResult As String
        '   strLead += r1.Item("QuotedSold") & " " & r1.Item("ParPrice") & vbCrLf
        Public QuotedSoldAmount As String
        '   strLead += r1.Item("ManagerNotes") & vbCrLf
        Public ManagerNotes As String
        '   strLead += r1.Item("Cash") & " " & r1.Item("Finance") & vbCrLf
        Public FinancingOption As String
        '       strLead += r1.Item("MarketingResults") & vbCrLf
        Public MarketingResults As String
        '   strLead += r1.Item("EmailAddress") & vbCrLf
        Public Cust_EmailAddress As String
        '   strLead += r1.Item("IspreviousCustomer") & " " & r1.Item("IsRecovery") & vbCrLf
        Public PreviousCustomer As String
    End Structure

    Public Exclu_ As Exclusions

    '' may  not need entire struct of employee
    '' may just need the name, email address and / or a telephone number 
    '' 
    '' methods: check rep can get email (function -> return boolean flag) 
    '' 

    '' Table Structure For Exclusions
    '' 
    '' Generated | Marketer | PLS | SLS | LastMResult | Phone 
    '' all boolean / bit fields 
    '' 


#Region "Constructor"
    Public Sub New()
 

        repsThatDontGetEmail = New ArrayList
        repsThatGetEmail = New ArrayList
        unassignedLeads = New ArrayList
        Exclu_ = New Exclusions
        arMessages = New ArrayList
        repEmailAddress = New ArrayList

    End Sub
#End Region

#Region "Get Exclusions"
    Public Function GetExclusions()
        Dim cnx_ex As SqlConnection = New SqlConnection(cnx_string)
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

#Region "Can the rep get email - OVERLOADED"
    ''
    '' overload for employeeID
    ''
    Public Function CanRepGetEmail(ByVal RepFName As String, ByVal RepLName As String)
        Dim can_get_email As Boolean = False
        Dim rep_cnx As SqlConnection = New SqlConnection(cnx_string)
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
        Dim rep_cnx As SqlConnection = New SqlConnection(cnx_string)
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

#Region "If the rep can get email - OVERLOADED"
    Public Function GetRepEmailAddress(ByVal RepFName As String, ByVal RepLName As String)
        Dim emailAddress As String
        Dim rep_cnx As SqlConnection = New SqlConnection(cnx_string)
        Dim cmdGet As SqlCommand = New SqlCommand("SELECT EmailAddress from tblTestEmployee where HasEmail = 1 and FName = '" & RepFName & "' and LName = '" & RepLName & "';", rep_cnx)
        rep_cnx.Open()
        emailAddress = cmdGet.ExecuteScalar
        rep_cnx.Close()
        rep_cnx = Nothing
        Return emailAddress
    End Function
    Public Function GetRepEmailAddress(ByVal EmployeeID As String)
        Dim emailAddress As String
        Dim rep_cnx As SqlConnection = New SqlConnection(cnx_string)
        Dim cmdGet As SqlCommand = New SqlCommand("SELECT EmailAddress from tblTestEmployee where HasEmail = 1 and EmployeeID = '" & EmployeeID & "';", rep_cnx)
        rep_cnx.Open()
        emailAddress = cmdGet.ExecuteScalar
        rep_cnx.Close()
        rep_cnx = Nothing
        Return emailAddress
    End Function



   
#End Region

#Region "Construct  Email - With Exclusions "
    Public Function ConstructMessageWithExclusions(ByVal RepFName As String, ByVal RepLName As String, ByVal LeadNum As String, ByVal _Exclusions As Exclusions, ByVal RepEmail As String)
        Dim ex_set As Exclusions = _Exclusions
        Dim eml_cnx As SqlConnection = New SqlConnection(cnx_string)
        Dim cmdGet As SqlCommand = New SqlCommand("SELECT * FROM EnterLead Where ID ='" & LeadNum & "';", eml_cnx)
        eml_cnx.Open()
        Dim r1 As SqlDataReader = cmdGet.ExecuteReader
        Dim strLead = ""
        strLead = "-------------" & vbCrLf
        Dim msgBody As EmailIssuedLeads.MessageBody
        While r1.Read
            '' Table Structure For Exclusions
            '' 
            Dim apptTimes() = Split(r1.Item("ApptTime"), " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            '' string 1/1/1900 11:00:00 AM
            ''     (0)           (1) 
            '' 
            Dim split2() = Split(apptTimes(1), ":", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim split3() = Split(apptTimes(1), " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            '' string 11:00:00 AM
            '' 11   00    00    AM
            '' (0)  (1)    (2)  

            '' string 11:00:00 AM
            ''           (0)    (1)
            Dim reconstituted As String = split2(0) & ":" & split2(2) & apptTimes(2)
            strLead += r1.Item("ApptDay") & " - " & r1.Item("ApptDate") & ", " & reconstituted & vbCrLf & vbCrLf
            '' Generated | Marketer | PLS | SLS | LastMResult | Phone 
            '' all boolean / bit fields 
            'strLead += r1.Item("ID") & vbCrLf
            strLead += "Customer ID : " & r1.Item("ID") & vbCrLf & vbCrLf
            'msgBody.RecordID = r1.Item("ID")
            strLead += r1.Item("Contact1LastName") & ", " & r1.Item("Contact1FirstName") & " & " & r1.Item("Contact2FirstName") & vbCrLf
            strLead += r1.Item("StAddress") & vbCrLf
            strLead += r1.Item("City") & ", " & r1.Item("State") & " " & r1.Item("Zip") & vbCrLf & vbCrLf

            Dim phone As String = r1.Item("HousePhone")
            If phone.ToString.Length >= 1 Then
                strLead += "House Phone: " & phone & vbCrLf
            ElseIf phone.ToString.Length <= 0 Then
                strLead = strLead
            End If

            '' phones 
            ' strLead += (r1.Item("HousePhone") & " " & r1.Item("Altphone1") & " " & r1.Item("AltPhone2")) & vbCrLf


            Dim alt1 As String = r1.Item("Altphone1")
            If alt1.ToString.Length >= 1 Then
                strLead += "Alt Phone1: " & alt1.ToString & vbCrLf
            ElseIf alt1.ToString.Length <= 0 Then
                strLead = strLead
            End If

            Dim alt2 As String = r1.Item("AltPhone2")
            If alt2.ToString.Length >= 1 Then
                strLead += "Alt Phone2: " & alt2.ToString & vbCrLf & vbCrLf
            ElseIf alt1.ToString.Length <= 0 Then
                strLead = strLead
            End If


            '' product 1 2, and 3 respectively 
            ' strLead += r1.Item("Product1") & " " & r1.Item("Product2") & " " & r1.Item("Product3")
            strLead += r1.Item("Product1") & vbCrLf

            Dim prod2 As String = r1.Item("Product2")
            If prod2.ToString.Length >= 1 Then
                strLead += r1.Item("Product2") & vbCrLf
            ElseIf prod2.ToString.Length <= 0 Then
                strLead += vbCrLf
            End If
            Dim prod3 As String = r1.Item("Product3")
            If prod3.ToString.Length >= 1 Then
                strLead += r1.Item("Product3") & vbCrLf
            ElseIf prod3.ToString.Length <= 0 Then
                strLead += vbCrLf
            End If

            strLead += "Special Instructions: " & r1.Item("SpecialInstruction") & vbCrLf & vbCrLf

            If ex_set.Marketer = True Then
                strLead = strLead
            ElseIf ex_set.Marketer = False Then
                strLead += "Lead Generated On: " & r1.Item("LeadGeneratedOn") & vbCrLf
            End If

            If ex_set.PLS = True Then
                strLead = strLead
            ElseIf ex_set.Marketer = False Then
                strLead += "Primary Lead Source: " & r1.Item("PrimaryLeadSource") & vbCrLf
            End If

            If ex_set.SLS = True Then
                strLead = strLead
            ElseIf ex_set.SLS = False Then
                strLead += "Secondary Lead Source: " & vbCrLf
            End If

            If ex_set.Marketer = True Then
                strLead = strLead
            ElseIf ex_set.Marketer = False Then
                strLead += "Marketer" & r1.Item("Marketer") & vbCrLf & vbCrLf
            End If

            If ex_set.LastMResult = True Then
                strLead = strLead
            ElseIf ex_set.LastMResult = False Then
                strLead += GetLastMarketingResult(LeadNum).ToString & vbCrLf
            End If

            strLead += "-------------"
            'If ex_set.Marketer = True Then
            '    strLead = strLead
            'strLead = strLead
            '    msgBody.Marketer = ""

            'ElseIf ex_set.Marketer = False Then
            '    strLead += r1.Item("Marketer") & vbCrLf

            '    msgBody.Marketer = r1.Item("Marketer")
            'End If

            'If ex_set.PLS = True Then
            '    strLead = strLead
            '    msgBody.PLS = ""
            'ElseIf ex_set.PLS = False Then
            '    strLead += r1.Item("PrimaryLeadSource") & vbCrLf
            '    msgBody.PLS = r1.Item("PrimaryLeadSource")
            'End If
            'If ex_set.SLS = True Then
            '    strLead = strLead
            '    msgBody.SLS = ""
            'ElseIf ex_set.SLS = False Then
            '    strLead += r1.Item("SecondaryLeadSource") & vbCrLf
            '    msgBody.SLS = r1.Item("SecondaryLeadSource")
            'End If
            'If ex_set.Generated = True Then
            '    strLead = strLead
            '    msgBody.LeadGeneratedOn = ""
            'ElseIf ex_set.Generated = False Then
            '    strLead += r1.Item("LeadGeneratedOn") & vbCrLf
            '    msgBody.LeadGeneratedOn = r1.Item("LeadGeneratedOn")
            'End If
            'strLead += (r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName")) & vbCrLf
            'msgBody.C1Name = (r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName"))
            'strLead += (r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName")) & vbCrLf
            'msgBody.C1Name = (r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName"))
            'strLead += (r1.Item("StAddress") & " " & r1.Item("City") & " " & r1.Item("state") & " " & r1.Item("Zip")) & vbCrLf
            'msgBody.Address = (r1.Item("StAddress") & " " & r1.Item("City") & " " & r1.Item("State") & " " & r1.Item("Zip"))
            'If ex_set.Phone = True Then
            '    strLead = strLead
            '    msgBody.PhoneNumbers = ""
            'ElseIf ex_set.Phone = False Then
            '    strLead += (r1.Item("HousePhone") & " " & r1.Item("Altphone1") & " " & r1.Item("AltPhone2")) & vbCrLf
            '    msgBody.PhoneNumbers = (r1.Item("HousePhone") & " | " & r1.Item("AltPhone1") & " | " & r1.Item("AltPhone2"))
            'End If
            'strLead += r1.Item("SpokeWith") & " " & r1.Item("Contact1WorkHours") & " " & r1.Item("Contact2WorkHours") & vbCrLf
            'msgBody.SpokeWithAndContactHours = (r1.Item("SpokeWith") & " " & r1.Item("Contact1WorkHours") & " " & r1.Item("Contact2WorkHours"))
            'strLead += r1.Item("Product1") & " " & r1.Item("Product2") & " " & r1.Item("Product3") & r1.Item("Color") & " " & r1.Item("ProductQTY") & vbCrLf
            'msgBody.Products = r1.Item("Product1").ToString & " | " & r1.Item("Product2").ToString & " | " & r1.Item("Product3").ToString
            'strLead += r1.Item("YearsOwned") & " " & r1.Item("HomeAge") & " " & r1.Item("HomeValue") & vbCrLf
            'msgBody.HomeInfo = r1.Item("YearsOwned") & " | " & r1.Item("HomeAge") & " | " & r1.Item("HomeValue")
            'strLead += r1.Item("ApptDate") & " " & r1.Item("ApptDay") & " " & r1.Item("ApptTime") & vbCrLf & vbCrLf
            'msgBody.ApptIfno = r1.Item("ApptDate") & " | " & r1.Item("ApptDay") & " | " & r1.Item("ApptTime")
            'strLead += r1.Item("SpecialInstruction") & vbCrLf & vbCrLf
            'msgBody.SpecialInstructions = r1.Item("SpecialInstruction")
            'strLead += r1.Item("Result") & vbCrLf
            'msgBody.LastResult = r1.Item("Result")
            'strLead += r1.Item("QuotedSold") & " " & r1.Item("ParPrice") & vbCrLf
            'msgBody.QuotedSoldAmount = r1.Item("QuotedSold") & " | " & r1.Item("ParPrice")
            'strLead += r1.Item("ManagerNotes") & vbCrLf
            'msgBody.ManagerNotes = r1.Item("ManagerNotes")
            'strLead += r1.Item("Cash") & " " & r1.Item("Finance") & vbCrLf
            'msgBody.FinancingOption = r1.Item("Cash") & " | " & r1.Item("Finance")

            'If ex_set.LastMResult = True Then
            '    strLead = strLead
            '    msgBody.MarketingResults = ""
            'ElseIf ex_set.LastMResult = False Then
            '    strLead += GetLastMarketingResult(LeadNum)
            '    msgBody.MarketingResults = GetLastMarketingResult(LeadNum)
            'End If
            'strLead += r1.Item("EmailAddress") & vbCrLf
            'msgBody.Cust_EmailAddress = r1.Item("EmailAddress")
            'strLead += r1.Item("IspreviousCustomer") & " " & r1.Item("IsRecovery") & vbCrLf
            'msgBody.PreviousCustomer = r1.Item("IsPreviousCustomer").ToString & " | " & r1.Item("IsRecovery").ToString
        End While
        r1.Close()
        eml_cnx.Close()
        'repsThatGetEmail.Add(RepFName & " " & RepLName)
        'repEmailAddress.Add(RepEmail)
        'arMessages.Add(msgBody)
        Return strLead
        'Return msgBody
    End Function
#Region "Bulk Mailers - Subs"

    Public Sub BulkMailWithExclusions(ByVal Replist As ArrayList, ByVal ApptDate As String)
        Dim arLeadNums As New ArrayList
        Dim g As Integer = 0
        For g = 0 To Replist.Count - 1
            Dim names() = Split(Replist(g), "|", -1, Microsoft.VisualBasic.CompareMethod.Text)
            '' repname  |  leadnum
            '' (0)          (1)
            Dim name As String = names(0)
            Dim leadnum As String = names(1)
            arLeadNums.Add(leadnum)
        Next

        ''
        '' 1 get distinct reps
        '' 2 get distint leadnums per unique rep
        '' 3 construct msg(s) per unique rep
        '' 4 concantenate msgs into one email per unique rep
        '' 5 group chronolgically, or sort by appt time asc (early to late)
        ''
        '' 1) select distinct(rep1) from enterlead where apptdate = whatever date order by appttime asc ==> list of distinct reps
        '' 2) select id from enterlead where apptdate = whatever date and and rep1 = rep from distinct list ==> list of lead nums per rep
        '' 3) send list to respective function for mail msg construction
        '' 4) self explanator add msgs to previous msgs
        '' 5) taken care of by step one with ASC flag in query
        '' 6) send 1(one) constructed msg to mailer to be mailed.
        '' 7) ???
        '' 8) Profit.

        Dim arListOfUniqueReps As New ArrayList
        Dim rep_cnx As SqlConnection = New SqlConnection(cnx_string)
        rep_cnx.Open()
        Dim cmdGet As SqlCommand = New SqlCommand("select Distinct(Rep1) from enterlead where apptdate = '" & ApptDate & "';", rep_cnx)
        Dim r1 As SqlDataReader = cmdGet.ExecuteReader
        While r1.Read
            If r1.Item("Rep1").ToString.Length >= 3 Then
                arListOfUniqueReps.Add(r1.Item("Rep1"))
            End If
        End While
        r1.Close()
        rep_cnx.Close()

        '' proxy array
        Dim prxF As New ArrayList
        Dim prxL As New ArrayList
        Dim i As Integer = 0
        For i = 0 To arListOfUniqueReps.Count - 1
            Dim names() = Split(arListOfUniqueReps(i), " ", -1)
            '' fname lname (0),(1)
            '' 
            Dim fname As String = names(0)
            Dim lname As String = names(1)
            prxF.Add(fname)
            prxL.Add(lname)
        Next

        '' now double check the unique rep can get email
        Dim ii As Integer = 0
        '' new list of reps that can / cannot get email
        Dim arScrubbed As New ArrayList

        For ii = 0 To prxF.Count - 1
            Dim canGetEmail As Boolean = CanRepGetEmail(prxF(ii), prxL(ii))
            If canGetEmail = True Then
                arScrubbed.Add(prxF(ii) & " " & prxL(ii))
            End If
        Next

        '' now we have a complete scrubbed unique rep list ==> arscrubbed
        '' 

        '' for each rep in unique rep list
        '' select ID,ApptDate,ApptTime from enterlead where ApptDate = '11/28/2014' and Rep1 = 'Radom Sales' order by ApptTime ASC;


        Dim aa As Integer = 0
        Dim listOfRepAndId(arScrubbed.Count - 1)
        Dim rep_cnx2 As SqlConnection = New SqlConnection(cnx_string)
        For aa = 0 To arScrubbed.Count - 1
            Dim listOfIDs As New ArrayList
            Dim cmdIDS As SqlCommand = New SqlCommand("SELECT ID,ApptDate,ApptTime FROM enterlead where ApptDate ='" & ApptDate & "' and Rep1 = '" & arScrubbed(aa) & "' and MarketingResults <> 'Called and Cancelled' order by ApptTime ASC;", rep_cnx2)
            rep_cnx2.Open()
            Dim r_ID As SqlDataReader = cmdIDS.ExecuteReader
            While r_ID.Read
                listOfIDs.Add(r_ID.Item("ID"))
            End While
            r_ID.Close()
            rep_cnx2.Close()
            listOfRepAndId(aa) = listOfIDs
        Next

        ''
        '' now we have a list of unique scrubbed reps and another array of arrays called listOfIDS per rep
        '' arScrubbed(x) = listofRepandID(x)
        ''
        ''  EACH     CONTAINS
        ''   OF       THIS
        ''  THESE     LIST
        '' |----|     |----|
        '' |rep | ==> |ID  |
        '' |----|     |----|
        ''            |ID  |
        ''            |----|


        '' now we need a list of email addresses per rep
        '' two birds with one stone here , split rep names and get email addresses.
        Dim prxyFAr2 As New ArrayList
        Dim prxyLAr2 As New ArrayList
        Dim arListOfEmails = New ArrayList
        Dim dd As Integer = 0
        For dd = 0 To arScrubbed.Count - 1
            Dim names3() = Split(arScrubbed(dd), " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim fname1 As String = names3(0)
            Dim lname1 As String = names3(1)
            '' mental note: this can be done better. doing the same thing 3 times now, but just to get the idea of out of my head, this is how i am building for now.
            '' 8-20-15 
            ''
            prxyFAr2.Add(fname1)
            prxyLAr2.Add(lname1)
            rep_cnx.Open()
            Dim cmdEMAIL As SqlCommand = New SqlCommand("SELECT EmailAddress from tblTestEmployee where FName = '" & fname1 & "' and LName = '" & lname1 & "';", rep_cnx)
            Dim EmailAddress As String = cmdEMAIL.ExecuteScalar
            rep_cnx.Close()
            arListOfEmails.add(EmailAddress)
        Next

        ''
        '' now we have arscrubbed,prxyFar2,prxyLar2,arlistofemails,arListofIDs ==> all scrubbed and specific to rep
        '' 

        '' now get exclusions set per rep
        '' so another array



        '' now we need to construct the ONE msg PER Rep
        '' 
        Dim ex_set As Exclusions = GetExclusions()

        Dim _CompleteMSG As String = ""
        Dim arMesssages As New ArrayList
        Dim msg_ As String = ""
        Dim msg2 As String = ""
        Dim bb As Integer = 0
        For bb = 0 To arScrubbed.Count - 1
            '' who are reps in list
            Dim names() = Split(arScrubbed(bb), " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim fname = names(0)
            Dim lname = names(1)
            Dim iteration As Integer = 0
            Dim email As String = arListOfEmails(bb)
            Dim arMsgs As ArrayList = listOfRepAndId(bb)
            Dim gg As Integer = 0
            For gg = 0 To arMsgs.Count - 1
                msg_ += ConstructMessageWithExclusions(fname, lname, arMsgs(gg), ex_set, email)
                msg2 += msg_ & vbCrLf
                msg_ = ""
            Next
            arMesssages.Add(msg2)
            msg2 = ""
        Next




        '' now mail them 
        '' 

        Dim ggg As Integer = 0
        For ggg = 0 To arMesssages.Count - 1
            Dim email As String = arListOfEmails(ggg)
            BulkMailWithExclusions(email, arMesssages(ggg))
            'MsgBox(arMesssages(ggg),MsgBoxStyle.Information ,"DEBUG INFO: WITH EXCLUSIONS")
        Next


    End Sub



    Public Sub BulkEmailWithoutExceptions(ByVal RepList As ArrayList, ByVal ApptDate As String)
        Dim arLeadNums As New ArrayList
        Dim g As Integer = 0
        For g = 0 To RepList.Count - 1
            Dim names() = Split(RepList(g), "|", -1, Microsoft.VisualBasic.CompareMethod.Text)
            '' repname  |  leadnum
            '' (0)          (1)
            Dim name As String = names(0)
            Dim leadnum As String = names(1)
            arLeadNums.Add(leadnum)
        Next

        ''
        '' 1 get distinct reps
        '' 2 get distint leadnums per unique rep
        '' 3 construct msg(s) per unique rep
        '' 4 concantenate msgs into one email per unique rep
        '' 5 group chronolgically, or sort by appt time asc (early to late)
        ''
        '' 1) select distinct(rep1) from enterlead where apptdate = whatever date order by appttime asc ==> list of distinct reps
        '' 2) select id from enterlead where apptdate = whatever date and and rep1 = rep from distinct list ==> list of lead nums per rep
        '' 3) send list to respective function for mail msg construction
        '' 4) self explanator add msgs to previous msgs
        '' 5) taken care of by step one with ASC flag in query
        '' 6) send 1(one) constructed msg to mailer to be mailed.
        '' 7) ???
        '' 8) Profit.

        Dim arListOfUniqueReps As New ArrayList
        Dim rep_cnx As SqlConnection = New SqlConnection(cnx_string)
        rep_cnx.Open()
        Dim cmdGet As SqlCommand = New SqlCommand("select Distinct(Rep1) from enterlead where apptdate = '" & ApptDate & "';", rep_cnx)
        Dim r1 As SqlDataReader = cmdGet.ExecuteReader
        While r1.Read
            arListOfUniqueReps.Add(r1.Item("Rep1"))
        End While
        r1.Close()
        rep_cnx.Close()

        '' proxy array
        Dim prxF As New ArrayList
        Dim prxL As New ArrayList
        Dim i As Integer = 0
        For i = 0 To arListOfUniqueReps.Count - 1
            If arListOfUniqueReps(i) = "" Then
            ElseIf arListOfUniqueReps(i) <> "" Then
                Dim names() = Split(arListOfUniqueReps(i), " ", -1)
                '' fname lname (0),(1)
                '' 
                Dim fname As String = names(0)
                Dim lname As String = names(1)
                prxF.Add(fname)
                prxL.Add(lname)
            End If
        Next

        '' now double check the unique rep can get email
        Dim ii As Integer = 0
        '' new list of reps that can / cannot get email
        Dim arScrubbed As New ArrayList

        For ii = 0 To prxF.Count - 1
            Dim canGetEmail As Boolean = CanRepGetEmail(prxF(ii), prxL(ii))
            If canGetEmail = True Then
                arScrubbed.Add(prxF(ii) & " " & prxL(ii))
            End If
        Next

        '' now we have a complete scrubbed unique rep list ==> arscrubbed
        '' 

        '' for each rep in unique rep list
        '' select ID,ApptDate,ApptTime from enterlead where ApptDate = '11/28/2014' and Rep1 = 'Radom Sales' order by ApptTime ASC;


        Dim aa As Integer = 0
        Dim listOfRepAndId(arScrubbed.Count - 1)
        Dim rep_cnx2 As SqlConnection = New SqlConnection(cnx_string)
        For aa = 0 To arScrubbed.Count - 1
            Dim listOfIDs As New ArrayList
            Dim cmdIDS As SqlCommand = New SqlCommand("SELECT ID,ApptDate,ApptTime FROM enterlead where ApptDate ='" & ApptDate & "' and Rep1 = '" & arScrubbed(aa) & "' and MarketingResults <> 'Called and Cancelled' order by ApptTime ASC;", rep_cnx2)
            rep_cnx2.Open()
            Dim r_ID As SqlDataReader = cmdIDS.ExecuteReader
            While r_ID.Read
                listOfIDs.Add(r_ID.Item("ID"))
            End While
            r_ID.Close()
            rep_cnx2.Close()
            listOfRepAndId(aa) = listOfIDs
        Next

        ''
        '' now we have a list of unique scrubbed reps and another array of arrays called listOfIDS per rep
        '' arScrubbed(x) = listofRepandID(x)
        ''
        ''  EACH     CONTAINS
        ''   OF       THIS
        ''  THESE     LIST
        '' |----|     |----|
        '' |rep | ==> |ID  |
        '' |----|     |----|
        ''            |ID  |
        ''            |----|


        '' now we need a list of email addresses per rep
        '' two birds with one stone here , split rep names and get email addresses.
        Dim prxyFAr2 As New ArrayList
        Dim prxyLAr2 As New ArrayList
        Dim arListOfEmails = New ArrayList
        Dim dd As Integer = 0
        For dd = 0 To arScrubbed.Count - 1
            Dim names3() = Split(arScrubbed(dd), " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim fname1 As String = names3(0)
            Dim lname1 As String = names3(1)
            '' mental note: this can be done better. doing the same thing 3 times now, but just to get the idea of out of my head, this is how i am building for now.
            '' 8-20-15 
            ''
            prxyFAr2.Add(fname1)
            prxyLAr2.Add(lname1)
            rep_cnx.Open()
            Dim cmdEMAIL As SqlCommand = New SqlCommand("SELECT EmailAddress from tblTestEmployee where FName = '" & fname1 & "' and LName = '" & lname1 & "';", rep_cnx)
            Dim EmailAddress As String = cmdEMAIL.ExecuteScalar
            rep_cnx.Close()
            arListOfEmails.add(EmailAddress)
        Next

        ''
        '' now we have arscrubbed,prxyFar2,prxyLar2,arlistofemails,arListofIDs ==> all scrubbed and specific to rep
        '' 

        '' now get exclusions set per rep
        '' so another array



        '' now we need to construct the ONE msg PER Rep
        '' 
        '' no exclusions needed on this sub 
        '' Dim ex_set As Exclusions = GetExclusions()

        Dim _CompleteMSG As String = ""
        Dim arMesssages As New ArrayList
        Dim msg_ As String = ""
        Dim msg2 As String = ""
        Dim bb As Integer = 0
        For bb = 0 To arScrubbed.Count - 1
            '' who are reps in list
            Dim names() = Split(arScrubbed(bb), " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim fname = names(0)
            Dim lname = names(1)
            Dim iteration As Integer = 0
            Dim email As String = arListOfEmails(bb)
            Dim arMsgs As ArrayList = listOfRepAndId(bb)
            Dim gg As Integer = 0
            For gg = 0 To arMsgs.Count - 1
                msg_ += ConstructMessageWithoutExclusions(fname, lname, arMsgs(gg), email)
                msg2 += msg_ & vbCrLf
                msg_ = ""
            Next
            arMesssages.Add(msg2)
            msg2 = ""
        Next




        '' now mail them 
        '' 

        Dim ggg As Integer = 0
        For ggg = 0 To arMesssages.Count - 1
            Dim email As String = arListOfEmails(ggg)
            BulkMailWithExclusions(email, arMesssages(ggg))
            'MsgBox(arMesssages(ggg),MsgBoxStyle.Information ,"DEBUG INFO: WITH EXCLUSIONS")
        Next
    End Sub

#End Region

    Public Function ConstructMessageWithoutExclusions(ByVal RepFName As String, ByVal RepLName As String, ByVal LeadNum As String, ByVal RepEmail As String)

        Dim eml_cnx As SqlConnection = New SqlConnection(cnx_string)
        Dim cmdGet As SqlCommand = New SqlCommand("SELECT * FROM EnterLead Where ID ='" & LeadNum & "';", eml_cnx)
        eml_cnx.Open()
        Dim msgBody As EmailIssuedLeads.MessageBody
        Dim r1 As SqlDataReader = cmdGet.ExecuteReader
        Dim strLead = ""
        strLead = "-------------" & vbCrLf
        While r1.Read
            ' '' Table Structure For Exclusions
            ' '' 
            ' '' Generated | Marketer | PLS | SLS | LastMResult | Phone 
            ' '' all boolean / bit fields 
            'strLead += r1.Item("ID") & vbCrLf
            'msgBody.RecordID = r1.Item("ID")
            'strLead += r1.Item("Marketer") & vbCrLf
            '    msgBody.Marketer = r1.Item("Marketer")
            'strLead += r1.Item("PrimaryLeadSource") & vbCrLf
            'msgBody.PLS = r1.Item("PrimaryLeadSource")
            'strLead += r1.Item("SecondaryLeadSource") & vbCrLf
            'msgBody.SLS = r1.Item("SecondaryLeadSource")
            'strLead += r1.Item("LeadGeneratedOn") & vbCrLf
            'msgBody.LeadGeneratedOn = r1.Item("LeadGeneratedOn")
            'strLead += (r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName")) & vbCrLf
            'msgBody.C1Name = (r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName"))
            'strLead += (r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName")) & vbCrLf
            'msgBody.C1Name = (r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName"))
            'strLead += (r1.Item("StAddress") & " " & r1.Item("City") & " " & r1.Item("state") & " " & r1.Item("Zip")) & vbCrLf
            'msgBody.Address = (r1.Item("StAddress") & " " & r1.Item("City") & " " & r1.Item("State") & " " & r1.Item("Zip"))
            'strLead += (r1.Item("HousePhone") & " " & r1.Item("Altphone1") & " " & r1.Item("AltPhone2")) & vbCrLf
            'msgBody.PhoneNumbers = (r1.Item("HousePhone") & " | " & r1.Item("AltPhone1") & " | " & r1.Item("AltPhone2"))
            'strLead += r1.Item("SpokeWith") & " " & r1.Item("Contact1WorkHours") & " " & r1.Item("Contact2WorkHours") & vbCrLf
            'msgBody.SpokeWithAndContactHours = (r1.Item("SpokeWith") & " " & r1.Item("Contact1WorkHours") & " " & r1.Item("Contact2WorkHours"))
            'strLead += r1.Item("Product1") & " " & r1.Item("Product2") & " " & r1.Item("Product3") & r1.Item("Color") & " " & r1.Item("ProductQTY") & vbCrLf
            'msgBody.Products = r1.Item("Product1") & " | " & r1.Item("Product2") & " | " & r1.Item("Product3")
            'strLead += r1.Item("YearsOwned") & " " & r1.Item("HomeAge") & " " & r1.Item("HomeValue") & vbCrLf
            'msgBody.HomeInfo = r1.Item("YearsOwned") & " | " & r1.Item("HomeAge") & " | " & r1.Item("HomeValue")
            'strLead += r1.Item("ApptDate") & " " & r1.Item("ApptDay") & " " & r1.Item("ApptTime") & vbCrLf & vbCrLf
            'msgBody.ApptIfno = r1.Item("ApptDate") & " | " & r1.Item("ApptDay") & " | " & r1.Item("ApptTime")
            'strLead += r1.Item("SpecialInstruction") & vbCrLf & vbCrLf
            'msgBody.SpecialInstructions = r1.Item("SpecialInstruction")
            'strLead += r1.Item("Result") & vbCrLf
            'msgBody.LastResult = r1.Item("Result")
            'strLead += r1.Item("QuotedSold") & " " & r1.Item("ParPrice") & vbCrLf
            'msgBody.QuotedSoldAmount = r1.Item("QuotedSold") & " | " & r1.Item("ParPrice")
            'strLead += r1.Item("ManagerNotes") & vbCrLf
            'msgBody.ManagerNotes = r1.Item("ManagerNotes")
            'strLead += r1.Item("Cash") & " " & r1.Item("Finance") & vbCrLf
            'msgBody.FinancingOption = r1.Item("Cash") & " | " & r1.Item("Finance")
            ' '' run stored proc here for marketing results
            'strLead += GetLastMarketingResult(LeadNum).ToString & vbCrLf
            'msgBody.MarketingResults = GetLastMarketingResult(LeadNum).ToString
            'strLead += r1.Item("EmailAddress") & vbCrLf
            'msgBody.Cust_EmailAddress = r1.Item("EmailAddress")
            'strLead += r1.Item("IspreviousCustomer") & " " & r1.Item("IsRecovery") & vbCrLf
            'msgBody.PreviousCustomer = r1.Item("IsPreviousCustomer") & " | " & r1.Item("IsRecovery")

            '' Table Structure For Exclusions
            '' 
            Dim apptTimes() = Split(r1.Item("ApptTime"), " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            '' string 1/1/1900 11:00:00 AM
            ''     (0)           (1)    (2)  
            '' 
            Dim split2() = Split(apptTimes(1), ":", -1, Microsoft.VisualBasic.CompareMethod.Text)
            '' string 11:00:00 AM
            '' 11   00    00 AM
            '' (0)  (1)    (2) 
            Dim reconstituted As String = split2(0) & ":" & split2(2) & apptTimes(2)
            strLead += r1.Item("ApptDay") & " - " & r1.Item("ApptDate") & ", " & reconstituted & vbCrLf & vbCrLf


            '' Generated | Marketer | PLS | SLS | LastMResult | Phone 
            '' all boolean / bit fields 
            'strLead += r1.Item("ID") & vbCrLf
            strLead += "Customer ID : " & r1.Item("ID") & vbCrLf & vbCrLf
            'msgBody.RecordID = r1.Item("ID")
            strLead += r1.Item("Contact1LastName") & ", " & r1.Item("Contact1FirstName") & " & " & r1.Item("Contact2FirstName") & vbCrLf
            strLead += r1.Item("StAddress") & vbCrLf
            strLead += r1.Item("City") & ", " & r1.Item("State") & " " & r1.Item("Zip") & vbCrLf & vbCrLf

            Dim phone As String = r1.Item("HousePhone")
            If phone.ToString.Length >= 1 Then
                strLead += "House Phone: " & phone & vbCrLf
            ElseIf phone.ToString.Length <= 0 Then
                strLead = ""
            End If

            '' phones 
            ' strLead += (r1.Item("HousePhone") & " " & r1.Item("Altphone1") & " " & r1.Item("AltPhone2")) & vbCrLf


            Dim alt1 As String = r1.Item("Altphone1")
            If alt1.ToString.Length >= 1 Then
                strLead += "Alt Phone1: " & alt1.ToString & vbCrLf
            ElseIf alt1.ToString.Length <= 0 Then
                strLead = strLead
            End If

            Dim alt2 As String = r1.Item("AltPhone2")
            If alt2.ToString.Length >= 1 Then
                strLead += "Alt Phone2: " & alt2.ToString & vbCrLf & vbCrLf
            ElseIf alt1.ToString.Length <= 0 Then
                strLead = strLead
            End If


            '' product 1 2, and 3 respectively 
            ' strLead += r1.Item("Product1") & " " & r1.Item("Product2") & " " & r1.Item("Product3")
            strLead += r1.Item("Product1") & vbCrLf

            Dim prod2 As String = r1.Item("Product2")
            If prod2.ToString.Length >= 1 Then
                strLead += r1.Item("Product2") & vbCrLf
            ElseIf prod2.ToString.Length <= 0 Then
                strLead += vbCrLf
            End If
            Dim prod3 As String = r1.Item("Product3")
            If prod3.ToString.Length >= 1 Then
                strLead += r1.Item("Product3") & vbCrLf
            ElseIf prod3.ToString.Length <= 0 Then
                strLead += vbCrLf
            End If

            strLead += "Special Instructions: " & r1.Item("SpecialInstruction") & vbCrLf & vbCrLf


            strLead += "Lead Generated On: " & r1.Item("LeadGeneratedOn") & vbCrLf



            strLead += "Primary Lead Source: " & r1.Item("PrimaryLeadSource") & vbCrLf




            strLead += "Secondary Lead Source: " & vbCrLf




            strLead += "Marketer" & r1.Item("Marketer") & vbCrLf & vbCrLf




            strLead += GetLastMarketingResult(LeadNum).ToString & vbCrLf


            strLead += "-------------"
        End While
        r1.Close()
        eml_cnx.Close()
        repsThatGetEmail.Add(RepFName & " " & RepLName)
        repEmailAddress.Add(RepEmail)
        arMessages.Add(msgBody)
        Return strLead
        'Return msgBody
    End Function

    Public Sub EMAIL_SINGLE_MarkupEmail_WITH_EXCLUSIONS(ByVal RepFName As String, ByVal RepLName As String, ByVal LeadNum As String, ByVal _Exclusions As Exclusions, ByVal RepEmail As String, ByVal Msg As String, ByVal Subject As String)
        Dim smptSERV As New SmtpClient
        Dim credentials As New Net.NetworkCredential
        credentials.Password = "bgfsreeffypxxxzr"
        credentials.UserName = "aaron.clay79@gmail.com"
        smptSERV.UseDefaultCredentials = False
        smptSERV.EnableSsl = True
        smptSERV.Port = 587 '' default for gmail ssl
        smptSERV.Host = "smtp.gmail.com"
        smptSERV.Credentials = credentials

        Dim eml_msg As New MailMessage("aaron.clay79@gmail.com", "aaron.clay79@gmail.com")
        Dim mailAddress As New MailAddress("aaron.clay79@gmail.com", "ImproveIt360!v2")

        With eml_msg
            .From = mailAddress
            .Subject = "Record ID: " & LeadNum.ToString
            .Body = Msg.ToString
            .IsBodyHtml = False
        End With
        smptSERV.Send(eml_msg)




    End Sub
    Public Sub BulkMailWithExclusions(ByVal RepEmail As String, ByVal MSG As String)
        Dim smptSERV As New SmtpClient
        Dim credentials As New Net.NetworkCredential
        credentials.Password = "bgfsreeffypxxxzr"
        credentials.UserName = "aaron.clay79@gmail.com"
        smptSERV.UseDefaultCredentials = False
        smptSERV.EnableSsl = True
        smptSERV.Port = 587 '' default for gmail ssl
        smptSERV.Host = "smtp.gmail.com"
        smptSERV.Credentials = credentials

        Dim eml_msg As New MailMessage(" aaron.clay79@gmail.com", "aaron.clay79@gmail.com") '' change the second part to actual rep's email for this to work right.
        Dim mailAddress As New MailAddress("aaron.clay79@gmail.com", "ImproveIt360!v2")

        With eml_msg
            .From = mailAddress
            .Subject = "Leads For The Day" & Date.Today.ToString
            .Body = MSG.ToString
            .IsBodyHtml = False
        End With
        smptSERV.Send(eml_msg)


    End Sub
    Public Sub BulkMailWithoutExclusions(ByVal RepEmail As String, ByVal MSG As String)
        Dim smptSERV As New SmtpClient
        Dim credentials As New Net.NetworkCredential
        credentials.Password = "bgfsreeffypxxxzr"
        credentials.UserName = "aaron.clay79@gmail.com"
        smptSERV.UseDefaultCredentials = False
        smptSERV.EnableSsl = True
        smptSERV.Port = 587 '' default for gmail ssl
        smptSERV.Host = "smtp.gmail.com"
        smptSERV.Credentials = credentials

        Dim eml_msg As New MailMessage(" aaron.clay79@gmail.com", "aaron.clay79@gmail.com") '' change the second part to actual rep's email for this to work right.
        Dim mailAddress As New MailAddress("aaron.clay79@gmail.com", "ImproveIt360!v2")

        With eml_msg
            .From = mailAddress
            .Subject = "Leads For The Day" & Date.Today.ToString
            .Body = MSG.ToString
            .IsBodyHtml = False
        End With
        smptSERV.Send(eml_msg)

    End Sub

    Public Sub EMAIL_SINGLE_MarkupEmail_WITHOUT_EXCLUSIONS(ByVal RepFName As String, ByVal RepLName As String, ByVal LeadNum As String, ByVal RepEmail As String, ByVal Msg As String, ByVal Subject As String)
        Dim smptSERV As New SmtpClient
        Dim credentials As New Net.NetworkCredential
        credentials.Password = "bgfsreeffypxxxzr"
        credentials.UserName = "aaron.clay79@gmail.com"
        smptSERV.UseDefaultCredentials = False
        smptSERV.EnableSsl = True
        smptSERV.Port = 587 '' default for gmail ssl
        smptSERV.Host = "smtp.gmail.com"
        smptSERV.Credentials = credentials

        Dim eml_msg As New MailMessage("aaron.clay79@gmail.com", "aaron.clay79@gmail.com")
        Dim mailAddress As New MailAddress("aaron.clay79@gmail.com", "ImproveIt360!v2")

        With eml_msg
            .From = mailAddress
            .Subject = "Record ID: " & LeadNum.ToString
            .Body = Msg.ToString
            .IsBodyHtml = False
        End With
        smptSERV.Send(eml_msg)


    End Sub

    Public Sub Send_BLAST_MAIL(ByVal _From As String, ByVal Recipient As String, ByVal _MSG As String, Subject As String)
        Dim smptSERV As New SmtpClient
        Dim credentials As New Net.NetworkCredential
        credentials.Password = "bgfsreeffypxxxzr"
        credentials.UserName = "aaron.clay79@gmail.com"
        smptSERV.UseDefaultCredentials = False
        smptSERV.EnableSsl = True
        smptSERV.Port = 587 '' default for gmail ssl
        smptSERV.Host = "smtp.gmail.com"
        smptSERV.Credentials = credentials

        Dim eml_msg As New MailMessage("aaron.clay79@gmail.com", Recipient)
        Dim mailAddress As New MailAddress("aaron.clay79@gmail.com", _From)

        With eml_msg
            .From = mailAddress
            .Subject = Subject
            .Body = _MSG
            .IsBodyHtml = False
        End With
        smptSERV.Send(eml_msg)

    End Sub



    Public Function GetLastMarketingResult(ByVal LeadNum As String)
        Dim lm_cnx As SqlConnection = New SqlConnection(cnx_string)
        lm_cnx.Open()
        Dim cmdGet As New SqlCommand("LastKnownMarketingResult", lm_cnx)
        cmdGet.Parameters.Add("@ID", SqlDbType.VarChar).Value = LeadNum
        cmdGet.CommandType = CommandType.StoredProcedure
        Dim r1 As SqlDataReader = cmdGet.ExecuteReader
        Dim m_result As String
        While r1.Read
            m_result = r1.Item("Mresult")
        End While
        r1.Close()
        lm_cnx.Close()
        lm_cnx = Nothing
        Return m_result
    End Function

    Public Function GetEnterLeadMarketingResult(ByVal LeadNum As String)
        Dim el_cnx As SqlConnection = New SqlConnection(cnx_string)
        el_cnx.Open()
        Dim cmdGet As SqlCommand = New SqlCommand("SELECT MarketingResults from EnterLead where ID = '" & LeadNum & "';", el_cnx)
        Dim res As String = cmdGet.ExecuteScalar
        el_cnx.Close()
        Dim CandC As Boolean
        If res = "Called and Cancelled" Then
            CandC = True
        ElseIf res <> "Called and Cancelled" Then
            CandC = False
        End If
        Return CandC

    End Function

#End Region

 

#Region "Email Marketing Manager List"
    '' select ID,appttime,Contact1FirstName,Contact2FirstName,Staddress,city,state,zip, product1,product2,product3,rep1,rep2 from enterlead
    ''where apptDate = '11/28/14' order by appttime asc;

    ''select FName,LName,EmailAddress from tblTestEmployee where department = 'Marketing Manager';

    Public Function GetMarketingManager()
        Dim mm As New MarketingManagerOBJ
        Dim mm_cnx As SqlConnection = New SqlConnection(cnx_string)
        mm_cnx.Open()
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT FName,LName,EmailAddress from tblTestEmployee WHERE Department = 'Marketing Manager';", mm_cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            mm.FName = r1.Item("FName")
            mm.LName = r1.Item("LName")
            mm.EmailAddress = r1.Item("EmailAddress")
        End While
        r1.Close()
        mm_cnx.Close()
        mm_cnx = Nothing
        Return mm
    End Function

    Public Function Generate_MarketingManager_List(ByVal ApptDate As String)
        Dim ArrayOfListOBJs As New ArrayList
        Dim mm_cnx As SqlConnection = New SqlConnection(cnx_string)
        Dim cmdGET As SqlCommand = New SqlCommand("select ID,appttime,Contact1FirstName,Contact2FirstName,Contact1LastName,Contact2LastName,Staddress,city,state,zip, product1,product2,product3,rep1,rep2,HousePhone,AltPhone1,AltPhone2 from EnterLead where apptDate = '" & ApptDate & "' order by appttime asc;", mm_cnx)
        mm_cnx.Open()
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim lead_obj As New MM_ListOBJ
            lead_obj.ID = r1.Item("ID")
            lead_obj.ApptTime = r1.Item("ApptTime")
            lead_obj.Contact1FirstName = r1.Item("Contact1FirstName")
            lead_obj.Contact2FirstName = r1.Item("Contact2FirstName")
            lead_obj.Contact1LastName = r1.Item("Contact1LastName")
            lead_obj.Contact2LastName = r1.Item("Contact2LastName")
            lead_obj.StAddress = r1.Item("StAddress")
            lead_obj.City = r1.Item("City")
            lead_obj.State = r1.Item("State")
            lead_obj.Zip = r1.Item("Zip")
            lead_obj.Product1 = r1.Item("Product1")
            lead_obj.Product2 = r1.Item("Product2")
            lead_obj.Product3 = r1.Item("Product3")
            lead_obj.Rep1 = r1.Item("Rep1")
            lead_obj.Rep2 = r1.Item("Rep2")
            lead_obj.Phone = r1.Item("HousePhone")
            lead_obj.AltPhone1 = r1.Item("AltPhone1")
            lead_obj.AltPhone2 = r1.Item("AltPhone2")
            ArrayOfListOBJs.Add(lead_obj)
            '' list of 'objs'
            '' 
        End While
        r1.Close()
        mm_cnx.Close()
        mm_cnx = Nothing
        Return ArrayOfListOBJs
    End Function


    Public Sub MailTheListToMarketingManager(ByVal listOfLeads As ArrayList, ByVal FName As String, ByVal LName As String, ByVal EmailAddress As String, ByVal ApptDate As String)
        Dim smptSERV As New SmtpClient
        Dim credentials As New Net.NetworkCredential
        credentials.Password = "bgfsreeffypxxxzr"
        credentials.UserName = "aaron.clay79@gmail.com"
        smptSERV.UseDefaultCredentials = False
        smptSERV.EnableSsl = True
        smptSERV.Port = 587 '' default for gmail ssl
        smptSERV.Host = "smtp.gmail.com"
        smptSERV.Credentials = credentials

        Dim daysplit() = Split(ApptDate, " ", -1)
        Dim _date As String = daysplit(0).ToString

        Dim eml_msg As New MailMessage("aaron.clay79@gmail.com", "aaron.clay79@gmail.com") '' change the to email address to email address variable
        Dim MM_Name As String = (FName & " " & LName)
        Dim mailAddress As New MailAddress("aaron.clay79@gmail.com", MM_Name)
        Dim dateTimeSplit() = Split(ApptDate, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim aptDate As String = dateTimeSplit(0)
        Dim g As Integer = 0
        Dim lines_ As String = "Issued Appointments For: " & Date.Today.DayOfWeek.ToString & ", " & aptDate
        lines_ += vbCrLf & vbCrLf

        For g = 0 To listOfLeads.Count - 1
            Dim y As MM_ListOBJ = listOfLeads(g)
            '' have to split off time

            Dim DateTime() = Split(y.ApptTime, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim time = DateTime(1)
            Dim timeSplit = Split(time, ":", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim HH As String = timeSplit(0).ToString
            Dim MM As String = timeSplit(1).ToString
            Dim _TimeConstructed = (HH & ":" & MM)
            Dim _AddressConstructed As String = y.StAddress & " " & y.City & ", " & y.State & " " & y.Zip
            
            lines_ += (y.ID & ":  " & _TimeConstructed & " | " & DetermineNames(y.Contact1FirstName, y.Contact1LastName, y.Contact2FirstName, y.Contact2LastName) & " | " & y.Phone & " | " & y.AltPhone1 & " | " & y.AltPhone2 & " | " & _AddressConstructed & " | " & y.Product1 & "-" & y.Product2 & "-" & y.Product3 & "-" & y.Rep1 & " " & y.Rep2) & vbCrLf & vbCrLf

        Next

        'MsgBox("DEBUG INFO: " & vbCrLf & lines_, MsgBoxStyle.Information, "DEBUG INFO: MM LIST CONSTRUCTION AND GATHERING")

        With eml_msg
            .From = mailAddress
            .Subject = "Issued Appointments For: " & Date.Today.DayOfWeek.ToString & ", " & aptDate
            .Body = lines_
            .IsBodyHtml = False
        End With
        smptSERV.Send(eml_msg)
    End Sub

    Private Function DetermineNames(ByVal c1Fname As String, ByVal c1lname As String, ByVal c2Fname As String, ByVal c2Lname As String)
        Dim name_Line As String = ""
        If c1lname = c2Lname Then
            name_Line = (c1Fname & " & " & c2Fname & " " & c1lname)
        ElseIf c2Lname = c1lname Then
            name_Line = (c1Fname & " & " & c2Fname & " " & c1lname)
        ElseIf c1lname <> c2Lname Then
            name_Line = (c1Fname & " " & c1lname & " and " & c2Fname & " " & c2Lname)
        ElseIf c2Lname <> c1lname Then
            name_Line = (c1Fname & " " & c1lname & " and " & c2Fname & " " & c2Lname)
        End If

        Return name_Line
    End Function

    Public Function IsCalledAndCancelled(ByVal LeadNum As String, ByVal ApptDate As String)
        Dim chk_cnx As SqlConnection = New SqlConnection(cnx_string)
        chk_cnx.Open()
        Dim cmdCNT As SqlCommand = New SqlCommand("SELECT count(ID) from enterlead where ApptDate = '" & ApptDate & "' and ID = '" & LeadNum & "' and MarketingResults <> 'Called and Cancelled'; ", chk_cnx)
        Dim cnt As Integer = cmdCNT.ExecuteScalar
        Dim ShouldMail As Boolean = False
        chk_cnx.Close()
        chk_cnx = Nothing
        If cnt >= 1 Then
            ShouldMail = True
        ElseIf cnt <= 0 Then
            ShouldMail = False
        End If
        Return ShouldMail
    End Function
#End Region

End Class
