Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient


Public Class convertLeadToStruct

    Private Const dev_cnx As String = "SERVER=PC-101\DEVMIRROREXPRESS;Database=devMirror;User Id=sa;Password=Legend1!"
    Private Const pro_cnx As String = "SERVER=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1"

    Public Structure EnterLead_Record
        Public RecID As String

        Public Marketer As String
        Public PLS As String
        Public SLS As String
        Public LeadGenOn As String
        Public C1FirstName As String
        Public C1LastName As String
        Public C2FirstName As String
        Public C2LastName As String
        Public StAddress As String

        Public City As String
        Public State As String
        Public Zip As String
        Public HousePhone As String
        Public AltPhone1 As String
        Public Phone1Type As String
        Public AltPhone2 As String
        Public Phone2Type As String
        Public SpokeWith As String

        Public C1WorkHours As String
        Public C2WorkHours As String
        Public Product1 As String
        Public Product2 As String
        Public Product3 As String
        Public Color As String
        Public ProductyQTY As String
        Public YearsOwned As String

        Public HomeAge As String
        Public ApptDate As String
        Public ApptDay As String
        Public ApptTime As String
        Public SpecialInstructions As String
        Public Lattitude As String
        Public Longitude As String
        Public TimeStampVal As String

        Public Rep1 As String
        Public Rep2 As String
        Public Result As String
        Public QuotedSold As String
        Public ParPrice As String
        Public Recoverable As String
        Public ManagerNotes As String
        Public Cash As String
        Public Finance As String
        Public P1QSSplit As String
        Public P2QSSplit As String
        Public P3QSSplit As String

        Public MarketingResults As String
        Public Confirmer As String
        Public DoNotMail As String
        Public DoNotCall As String
        Public EmailAddress As String
        Public Product1Acro As String
        Public Product2Acro As String
        Public Product3Acro As String
        Public MarketingManager As String
        Public SalesManager As String
        Public KillPending As String
        Public IssueNotes As String

        Public NeedsSalesResults As String
        Public SetNotes As String
        Public IsPreviousCustomer As String
        Public IsRecovery As String
        Public LastUpdated As String

    End Structure

    '' TRUE = DEV | FALSE = Production
    '' 

    Public Function ConvertToStructure(ByVal RecID As String, ByVal Dev_Or_Pro As Boolean)
        Dim Lead As EnterLead_Record
        Dim cnx As SqlConnection
        If Dev_Or_Pro = True Then
            cnx = New SqlConnection(dev_cnx)
        ElseIf Dev_Or_Pro = False Then
            cnx = New SqlConnection(pro_cnx)
        End If
        cnx.Open()
        Dim strGET As String = "SELECT * FROM EnterLead WHERE ID = '" & RecID & "';"
        Dim cmdGET As New SqlCommand(strGET, cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Lead.RecID = ConvertIfNull(r1.Item("ID"))
            Lead.Marketer = ConvertIfNull(r1.Item("Marketer"))
            Lead.PLS = ConvertIfNull(r1.Item("PrimaryLeadSource"))
            Lead.SLS = ConvertIfNull(r1.Item("SecondaryLeadSource"))
            Lead.LeadGenOn = ConvertIfNull(r1.Item("LeadGeneratedOn"))
            Lead.C1FirstName = ConvertIfNull(r1.Item("Contact1FirstName"))
            Lead.C1LastName = ConvertIfNull(r1.Item("Contact1LastName"))
            Lead.C2FirstName = ConvertIfNull(r1.Item("Contact2FirstName"))
            Lead.C2LastName = ConvertIfNull(r1.Item("Contact2LastName"))
            Lead.StAddress = ConvertIfNull(r1.Item("StAddress"))
            Lead.City = ConvertIfNull(r1.Item("City"))
            Lead.State = ConvertIfNull(r1.Item("State"))
            Lead.Zip = ConvertIfNull(r1.Item("Zip"))
            Lead.HousePhone = ConvertIfNull(r1.Item("HousePhone"))
            Lead.AltPhone1 = ConvertIfNull(r1.Item("AltPhone1"))
            Lead.Phone1Type = ConvertIfNull(r1.Item("Phone1Type"))
            Lead.AltPhone2 = ConvertIfNull(r1.Item("AltPhone2"))
            Lead.Phone2Type = ConvertIfNull(r1.Item("Phone2Type"))
            Lead.SpokeWith = ConvertIfNull(r1.Item("SpokeWith"))
            Lead.C1WorkHours = ConvertIfNull(r1.Item("Contact1WorkHours"))
            Lead.C2WorkHours = ConvertIfNull(r1.Item("Contact2WorkHours"))
            Lead.Product1 = ConvertIfNull(r1.Item("Product1"))
            Lead.Product2 = ConvertIfNull(r1.Item("Product2"))
            Lead.Product3 = ConvertIfNull(r1.Item("Product3"))
            Lead.Color = ConvertIfNull(r1.Item("Color"))
            Lead.ProductyQTY = ConvertIfNull(r1.Item("ProductQTY"))
            Lead.YearsOwned = ConvertIfNull(r1.Item("YearsOwned"))
            Lead.HomeAge = ConvertIfNull(r1.Item("HomeAge"))
            Lead.ApptDate = ConvertIfNull(r1.Item("ApptDate"))
            Lead.ApptTime = ConvertIfNull(r1.Item("ApptTime"))
            Lead.ApptDay = ConvertIfNull(r1.Item("ApptDay"))
            Lead.SpecialInstructions = ConvertIfNull(r1.Item("SpecialInstruction"))
            Lead.TimeStampVal = ConvertIfNull(r1.Item("TimeStampVal"))
            Lead.Rep1 = ConvertIfNull(r1.Item("Rep1"))
            Lead.Rep2 = ConvertIfNull(r1.Item("Rep2"))
            Lead.Result = ConvertIfNull(r1.Item("Result"))
            Lead.QuotedSold = ConvertIfNull(r1.Item("QuotedSold"))
            Lead.ParPrice = ConvertIfNull(r1.Item("ParPrice"))
            Lead.Recoverable = ConvertIfNull(r1.Item("Recoverable"))
            Lead.ManagerNotes = ConvertIfNull(r1.Item("ManagerNotes"))
            Lead.Cash = ConvertIfNull(r1.Item("Cash"))
            Lead.Finance = ConvertIfNull(r1.Item("Finance"))
            Lead.P1QSSplit = ConvertIfNull(r1.Item("P1QSSplit"))
            Lead.P2QSSplit = ConvertIfNull(r1.Item("P2QSSplit"))
            Lead.P3QSSplit = ConvertIfNull(r1.Item("P3QSSplit"))
            Lead.MarketingResults = ConvertIfNull(r1.Item("MarketingResults"))
            Lead.Confirmer = ConvertIfNull(r1.Item("Confirmer"))
            Lead.DoNotCall = ConvertIfNull(r1.Item("DoNotCall"))
            Lead.DoNotMail = ConvertIfNull(r1.Item("DoNotMail"))
            Lead.EmailAddress = ConvertIfNull(r1.Item("EmailAddress"))
            Lead.Product1Acro = ConvertIfNull(r1.Item("Productacro1"))
            Lead.Product2Acro = ConvertIfNull(r1.Item("Productacro2"))
            Lead.Product3Acro = ConvertIfNull(r1.Item("Productacro3"))
            Lead.MarketingManager = ConvertIfNull(r1.Item("MarketingManager"))
            Lead.SalesManager = ConvertIfNull(r1.Item("SalesManager"))
            Lead.KillPending = ConvertIfNull(r1.Item("KillPending"))
            Lead.IssueNotes = ConvertIfNull(r1.Item("IssueNotes"))
            Lead.NeedsSalesResults = ConvertIfNull(r1.Item("NeedsSaleResult"))
            Lead.SetNotes = ConvertIfNull(r1.Item("SetNotes"))
            Lead.IsPreviousCustomer = ConvertIfNull(r1.Item("IsPreviousCustomer"))
            Lead.IsRecovery = ConvertIfNull(r1.Item("IsRecovery"))
            Lead.LastUpdated = ConvertIfNull(r1.Item("LastUpdated"))
        End While
        cnx.Close()
        cnx = Nothing
        Return Lead
    End Function

    Private Function ConvertIfNull(ByVal Data As Object)
        Dim str As String = ""

        If IsDBNull(Data) = True Then
            str = ""
        ElseIf IsDBNull(Data) = False Then
            str = Data.ToString
        End If

        Return str
    End Function
End Class
