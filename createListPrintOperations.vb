Imports System.Text
Imports System.IO
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient


Public Class createListPrintOperations

    Public Structure LeadToPrint ''DEFAULT
        Public ID As String
        Public FName As String
        Public LName As String
        Public FName2 As String
        Public LName2 As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
        Public Phone1 As String
        Public Phone2 As String
        Public Phone3 As String
        Public Prod1 As String
        Public Prod2 As String
        Public Prod3 As String
        Public LeadGeneratedOn As String
        Public ApptDate As String
        Public ApptTime As String
        Public ApptDay As String
        Public Result As String
        Public MarketingResults As String
        Public IsPreviousCustomer As String
        Public IsRecovery As String
    End Structure

    Public Structure PCInfoToPrint '' in the instance that a previous sale exists for LeadNumber this is the additional info that needs to be pulled 
        Public recID As String
        Public LeadNum As String
        Public LeadHistoryID As String
        Public JobCloseDate As String
        Public ContractAmount As String
        Public Cash As Boolean
        Public Finance As Boolean
        Public Product1 As String
        Public Product2 As String
        Public Product3 As String
    End Structure


#Region "notes"
    ''10-7-2015
    ''
    '' 5 structures needed
    ''
    '' DEFAULT = Select Distinct(Enterlead.ID), Contact1LastName, Contact1FirstName, Contact2LastName, Contact2FirstName, StAddress, City, State, Zip, HousePhone, AltPhone1, AltPhone2, Enterlead.Product1, Enterlead.Product2, Enterlead.Product3, LeadGeneratedOn, ApptDate, ApptDay, ApptTime, Result, MarketingResults, IsPreviousCustomer, IsRecovery

    '' 1) for DEFAULT (no order by selected)
    ''    a) no extra info needed
    ''    i) DEFAULT OBJ

    '' 2) ORDER BY: Reference Rating 
    ''    b) col: ReferenceRating
    ''    ii) DEFAULT + col

    '' 3) ORDER BY: Sale Closed Date
    ''    c) col: JobClosed
    ''    iii) DEFAULT + col

    '' 4) ORDER BY: Loan Satisfied Date
    ''    d) col: ExpectedPayOff
    ''   iv) DEFAULT + col

    '' 5) ORDER BY: Approved Load Amount
    ''    e) col: ApprovedFor
    ''    v) DEFAULT + col

    '' idea 1 
    '' Use a Hash Table For Key:Value Pairs  || DONT KNOW IF YOU CAN STORE an arraylist as a 'value' in hash table
    '' MSDN Ref: https://msdn.microsoft.com/en-us/library/system.collections.hashtable(v=vs.110).aspx
    '' MSDN Ref: https://msdn.microsoft.com/en-us/library/system.collections.hashtable.add(v=vs.110).aspx   -> Outlining key and value as OBJECTS
    '' Framework 4 Supports this Method
    '' according to MSDN,    ADD(object,object)  method can use 'object' type when storing either keys or values.
    '' 
    '' 
    '' example  entry1 = Key "List(of type default)" : value = arLstOfDefaultObjs as arraylist housing DEFAULT Structs
    ''          entry2 = Key "List(of type ReferenceRating)" : value = arLstOfRefRating as arraylist housing RefRating Structs
    '' 

    '' When pulling out the 'value' from hastable, should be able to use 'if typeOf ' for a decission struture / select case to know what type of object to expect from return / expression
    '' 

    '' idea 2 || Lengthy work , has to be a better way.
    '' Write a seperate Sub for each type then, depending on what was selected for ORDER BY use respective Sub
    '' 

    '' idea3 || Generic defined Multi-Dim array  arObjects(0,0)   first ordinal = description  ie: 'list of reference rating type' | second ordinal = List(of ReferenceRating) 
    '' 
    '' 

    '' 10-8-2015
    '' hash table idea works, but after more consideration:
    '' all irrelevant ideas. The fields are always the same as defined in the query the only thing that is really changing is the criteria which we don't have to render.
    '' just pass the (slct) string to a method to run query, grab results, then render. 
    '' So, just one structure : DEFAULT
    '' 

#End Region




    Public arListOfLeadsToDisplay As List(Of LeadToPrint)

    Private Const sql_cnx As String = "SERVER=192.168.1.2;DATABASE=ISS;User Id=sa;Password=spoken1;"
    Private Const Test_Directory As String = "C:\Users\Clay\Desktop\Print Leads\"
    Private Const Production_Directory As String = "\\server.greenworks.local\Company\ISS\Print Leads\"

    Public Sub New()
        arListOfLeadsToDisplay = New List(Of LeadToPrint)
    End Sub

    Public Sub CreateWireFrameHTML(ByVal QueryString As String)

        arListOfLeadsToDisplay = GrabRowsForHTML(QueryString)

        Dim guid As Guid = guid.NewGuid
        Dim dateTime As String = Date.Now
        Dim ext As String = ".htm"

        'Dim style As String = "<style type='text/css'>.elem1{font-family:""Verdana"",sans-serif;font-size:1.2em;color:blue;}.fancyTable{position:fixed;left:2em;top:3em;padding:2px;margin:3px;text-align:center;"
        'style = style + "border-right-style:solid;border-right-width:1px;border-right-color:black;box-shadow:2px 2px 2px grey;border-collapse: collapse;}"
        'style = style + "th{padding:3px;border-bottom-style:solid;border-bottom-color:black;border-bottom-width:1px;border-right-style:solid;border-right-color:grey;border-right-width:1px;"
        'style = style + "background-color:hsl(30,75%,74%);color:hsl(30,100%,33%);font-family:""Verdana"",sans-serif;font-size:1.1em;}"
        'style = style + "td{padding:3px;font-family:""Verdana"",sans-serif;font-size:0.7em;}"
        'style = style + "tr:nth-child(even){background-color:hsl(30,100%,94%);}tr:nth-child(odd){background-color:hsl(30,100%,80%);}"
        'style = style + "</style>"

        Dim title As String = "<title>Create List For " & Date.Today.ToShortDateString & "</title>"

        Dim style As String = "<style type='text/css'>" & vbCrLf
        style += "span.data {font-family:""Verdana"",sans-serif;font-size:1.3em;color:black;}" & vbCrLf
        style += ".elem1 {font-family:""Verdana"",sans-serif;color:blue;font-size:1.3em;}" & vbCrLf
        style += ".tbl {font-family:""Verdana"",sans-serif;color:black;border-collapse:collapse;margin:5px;box-shadow:2px 2px 2px grey;}" & vbCrLf
        style += "tr {padding:3px;}" & vbCrLf
        style += "tr:nth-child(odd) {background-color:white;}" & vbCrLf
        style += "tr:nth-child(even) {background-color:hsl(0,0%,90%);}" & vbCrLf
        style += "th {font-family:""Verdana"",sans-serif;font-size:1.1em;color:black;padding:3px;text-align:center;background-color:hsl(0,0%,80%);}" & vbCrLf
        style += "td {font-family:""Verdana"",sans-serif;font-size:0.8em;padding:7px;text-align:center;}" & vbCrLf
        style += "tr.sale {font-family:""Verdana"",sans-serif;font-weight:normal;font-size:1em;background-color:hsl(120,100%,75%);}" & vbCrLf
        style += "@media print { span.data {display:none;} tr:nth-child(even){background-color:white;}tr:nth-child(odd){background-color:white;}" & vbCrLf
        style += "tr.sale{background-color:white;font-size:0.8em;}th{font-size:0.7em;}td{font-size:0.5em;}.tbl{box-shadow:0px 0px 0px white;} @page{size:landscape;}}" & vbCrLf
        style += "</style>"


        Dim jScript As String = "<script type='text/javascript'></script>"

        Dim streamWriter As New StreamWriter(Production_Directory & guid.ToString & ext)

        streamWriter.WriteLine("<!DOCTYPE html>")
        streamWriter.WriteLine("<head>")
        streamWriter.WriteLine("" & title & "")
        streamWriter.WriteLine("" & style & "")
        streamWriter.WriteLine("" & jScript & "")
        streamWriter.WriteLine("</head>")
        streamWriter.WriteLine("<body>")
        streamWriter.WriteLine("<span id='elem1' class='elem1'>Customer List For: " & Date.Today.ToShortDateString & "</span><br /><br /><br /><span>" & QueryString.ToString & "</span>")
        streamWriter.WriteLine("<br /><br />")
        streamWriter.WriteLine("<table id='tbDat' class='tbl'>")
        streamWriter.WriteLine("<tr><th></th><th>ID</th><th>Name</th><th>Address</th><th>Phone</th><th>Product(s)</th><th>Generated On</th><th>Appt. Date</th><th>Appt. Day</th><th>Appt. Time</th><th>Result</th></tr>")

        Dim z As LeadToPrint
        For Each z In arListOfLeadsToDisplay
            Dim timeSplit() = Split(z.ApptTime, " ", 2)
            Dim apptTime As String = timeSplit(1).ToString
            Dim LeadGenSplit() = Split(z.LeadGeneratedOn, " ", 2)

            Dim leadGen As String = LeadGenSplit(0)
            Dim recID As String = z.ID
            Dim prodString = (z.Prod1 & "<br />" & z.Prod2 & "<br />" & z.Prod3)
            Dim PhoneString = (z.Phone1 & "<br />" & z.Phone2 & "<br />" & z.Phone3)
            Dim AddyString = (z.StAddress & "<br />" & z.City & ", " & z.State & " " & z.Zip)


            ''10-8-2015 8:20am
            '' dont' worry about icons just yet
            '' double check methods are working correctly.

            ''10-8-2015 9:45am
            '' split for icons per result
            '' 

            Select Case z.Result
                Case Is = "Sale"
                    streamWriter.WriteLine("<tr><td><img src='IMGS\sale28x28.jpg' alt='Sale' title='Sale'/><td>" & z.ID & "</td><td>" & z.FName & " " & z.LName & "<br />" & z.FName2 & " " & z.LName2 & "</td><td>" & AddyString & "</td><td>" & PhoneString & "</td><td>" & prodString & "</td><td>" & leadGen & "</td><td>" & z.ApptDate & "</td><td>" & z.ApptDay & "</td><td>" & apptTime & "</td><td>" & z.Result & "</td></tr>")
                    '' grab sales cross query
                    Dim g As New List(Of PCInfoToPrint)
                    g = GrabSaleInfo(z.ID)
                    Dim a As PCInfoToPrint
                    Dim cnt As Integer = 0
                    For Each a In g
                        cnt += 1
                        Dim p_String As String = (a.Product1 & "<br />" & a.Product2 & "<br />" & a.Product3)
                        If a.Cash = True Then
                            Dim f_currency As String = FormatCurrency(a.ContractAmount, 2, Microsoft.VisualBasic.TriState.UseDefault, Microsoft.VisualBasic.TriState.UseDefault, Microsoft.VisualBasic.TriState.UseDefault)
                            streamWriter.WriteLine("<tr class='sale'><td/><td /><td >Job " & cnt & "</td><td>" & p_String & "</td><td>Job Closed<br />" & a.JobCloseDate & "</td><td>Contract Amt<br />" & f_currency & "</td><td>Cash</td><td/><td/><td/><td/></tr>")
                        ElseIf a.Finance = True Then
                            Dim f_currency As String = FormatCurrency(a.ContractAmount, 2, Microsoft.VisualBasic.TriState.UseDefault, Microsoft.VisualBasic.TriState.UseDefault, Microsoft.VisualBasic.TriState.UseDefault)
                            streamWriter.WriteLine("<tr class='sale'><td/><td /><td >Job " & cnt & "</td><td>" & p_String & "</td><td>Job Closed<br />" & a.JobCloseDate & "</td><td>Contract Amt<br />" & f_currency & "</td><td>Finance</td><td/><td/><td/><td/></tr>")
                        End If
                    Next
                    Exit Select
                Case Is = "Demo/No Sale"
                    streamWriter.WriteLine("<tr><td><img src='IMGS\DemoNoSale28x28.jpg' alt='Demo/No Sale' title='Demo/No Sale'/><td>" & z.ID & "</td><td>" & z.FName & " " & z.LName & "<br />" & z.FName2 & " " & z.LName2 & "</td><td>" & AddyString & "</td><td>" & PhoneString & "</td><td>" & prodString & "</td><td>" & leadGen & "</td><td>" & z.ApptDate & "</td><td>" & z.ApptDay & "</td><td>" & apptTime & "</td><td>" & z.Result & "</td></tr>")
                    Exit Select
                Case Is = "Reset"
                    streamWriter.WriteLine("<tr><td><img src='IMGS\reset28x28.jpg' alt='Reset' title='Reset'/><td>" & z.ID & "</td><td>" & z.FName & " " & z.LName & "<br />" & z.FName2 & " " & z.LName2 & "</td><td>" & AddyString & "</td><td>" & PhoneString & "</td><td>" & prodString & "</td><td>" & leadGen & "</td><td>" & z.ApptDate & "</td><td>" & z.ApptDay & "</td><td>" & apptTime & "</td><td>" & z.Result & "</td></tr>")
                    Exit Select
                Case Else
                    streamWriter.WriteLine("<tr><td /><td>" & z.ID & "</td><td>" & z.FName & " " & z.LName & "<br />" & z.FName2 & " " & z.LName2 & "</td><td>" & AddyString & "</td><td>" & PhoneString & "</td><td>" & prodString & "</td><td>" & leadGen & "</td><td>" & z.ApptDate & "</td><td>" & z.ApptDay & "</td><td>" & apptTime & "</td><td>" & z.Result & "</td></tr>")
                    Exit Select
            End Select
            


            '' mockup example only format
            'streamWriter.WriteLine("<tr><td>" & z.ID & "</td><td>" & z.FName & "</td><td>" & z.LName & "</td><td>" & z.StAddress & " " & z.City & " " & z.State & " " & z.Zip & "</td><td>" & z.Phone1 & "</td><td>" & z.Phone2 & "</td><td>" & z.Prod1 & "</td><td>" & z.Prod2 & "</td><td>" & z.Prod3 & "</td><td>" & z.ApptDate & "</td><td>" & timeSplit(1).ToString & "</td></tr>")

            '' printLeadExample.xls Format
            '' 10-6-2015


        Next

        streamWriter.WriteLine("</table></body></html>")
        streamWriter.Flush()
        streamWriter.Close()

        System.Diagnostics.Process.Start(Production_Directory & guid.ToString & ext)


    End Sub

    Private Function GrabRowsForHTML(ByVal Query_String As String)
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim arResults As New List(Of LeadToPrint)
        Dim cmdGET As New SqlCommand(Query_String, cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        '' DEFAULT = Select Distinct(Enterlead.ID), Contact1LastName, Contact1FirstName, Contact2LastName, Contact2FirstName, StAddress, City, State, Zip, 
        '' HousePhone, AltPhone1, AltPhone2, EnterLead.Product1, EnterLead.Product2, EnterLead.Product3, 
        '' LeadGeneratedOn, ApptDate, ApptDay, ApptTime, Result, MarketingResults, IsPreviousCustomer, IsRecovery
        While r1.Read
            Dim y As New LeadToPrint
            y.ID = r1.Item("ID")
            y.FName = r1.Item("Contact1FirstName")
            y.LName = r1.Item("Contact1LastName")
            y.FName2 = r1.Item("Contact2FirstName")
            y.LName2 = r1.Item("Contact2LastName")
            y.StAddress = r1.Item("StAddress")
            y.City = r1.Item("City")
            y.State = r1.Item("State")
            y.Zip = r1.Item("Zip")
            y.Phone1 = r1.Item("HousePhone")
            y.Phone2 = r1.Item("AltPhone1")
            y.Phone3 = r1.Item("AltPhone2")
            y.Prod1 = r1.Item("Product1")
            y.Prod2 = r1.Item("Product2")
            y.Prod3 = r1.Item("Product3")
            y.LeadGeneratedOn = r1.Item("LeadGeneratedOn")
            y.ApptDate = r1.Item("ApptDate")
            y.ApptTime = r1.Item("ApptTime")
            y.ApptDay = r1.Item("ApptDay")
            y.Result = r1.Item("Result")
            y.MarketingResults = r1.Item("MarketingResults")
            y.IsPreviousCustomer = r1.Item("IsPreviousCustomer")
            y.IsRecovery = r1.Item("IsRecovery")
            arResults.Add(y)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arResults
    End Function

    Private Function GrabSaleInfo(ByVal LeadNum As String)
        Dim cnxS As New SqlConnection(sql_cnx)
        cnxS.Open()
        Dim cmdGET As New SqlCommand("SELECT ID,LeadNum,LeadHistoryID,Product1,Product2,Product3,JobClosed,ContractAmount,Cash,Finance FROM SaleDetail WHERE LeadNum = '" & LeadNum & "';", cnxS)
        Dim arResults As New List(Of PCInfoToPrint)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim y As New PCInfoToPrint
            y.recID = r1.Item("ID")
            y.LeadNum = r1.Item("LeadNum")
            y.LeadHistoryID = r1.Item("LeadHistoryID")
            y.Product1 = r1.Item("Product1")
            y.Product2 = r1.Item("Product2")
            y.Product3 = r1.Item("Product3")
            y.JobCloseDate = r1.Item("JobClosed")
            y.ContractAmount = r1.Item("ContractAmount")
            y.Cash = r1.Item("Cash")
            y.Finance = r1.Item("Finance")
            arResults.Add(y)
        End While
        r1.Close()
        cnxS.Close()
        cnxS = Nothing
        Return arResults
    End Function


    Public Function TestList()
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("SELECT TOP(10)ID FROM EnterLead;", cnx)
        Dim arLeadNum As New ArrayList
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            arLeadNum.Add(r1.Item("ID"))
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arLeadNum
    End Function


End Class
