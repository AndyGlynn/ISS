Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports System.IO
Imports System.Text


Public Class CompanyInformation
    '' sql connection string
    '' 
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)

    '' private variables for properties
    '' 
    Private StNum As String = ""
    Private StAddy As String = ""
    Private Addy2 As String = ""
    Private Cty As String = ""
    Private St As String = ""
    Private Zp As String = ""
    Private Logo_Dir As String = ""
    Private ContactPhNum As String = ""
    Private ContactFxNum As String = ""
    Private CompWebAddy As String = ""
    Private CompName As String = ""

    '' public properties 
    '' 
    Public Property StNumber() As String
        Get
            Return StNum
        End Get
        Set(ByVal value As String)
            StNum = value
        End Set
    End Property
    Public Property StAddress() As String
        Get
            Return StAddy
        End Get
        Set(ByVal value As String)
            StAddy = value
        End Set
    End Property
    Public Property Address_Line_2() As String
        Get
            Return Addy2
        End Get
        Set(ByVal value As String)
            Addy2 = value
        End Set
    End Property
    Public Property City() As String
        Get
            Return Cty
        End Get
        Set(ByVal value As String)
            Cty = value
        End Set
    End Property
    Public Property State() As String
        Get
            Return St
        End Get
        Set(ByVal value As String)
            St = value
        End Set
    End Property
    Public Property Zip() As String
        Get
            Return Zp
        End Get
        Set(ByVal value As String)
            Zp = value
        End Set
    End Property
    Public Property Logo_Directory() As String
        Get
            Return Logo_Dir
        End Get
        Set(ByVal value As String)
            Logo_Dir = value
        End Set
    End Property
    Public Property ContactPhoneNumber() As String
        Get
            Return ContactPhNum
        End Get
        Set(ByVal value As String)
            ContactPhNum = value
        End Set
    End Property
    Public Property ContactFaxNumber() As String
        Get
            Return ContactFxNum
        End Get
        Set(ByVal value As String)
            ContactFxNum = value
        End Set
    End Property
    Public Property Company_WebSite() As String
        Get
            Return CompWebAddy
        End Get
        Set(ByVal value As String)
            CompWebAddy = value
        End Set
    End Property
    Public Property Company_Name() As String
        Get
            Return CompName
        End Get
        Set(ByVal value As String)
            CompName = value
        End Set
    End Property

    Public Sub GetInformation()
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("select * from iss.dbo.companyinfo", cnn)
            Dim r1 As SqlDataReader
            cnn.Open()
            r1 = cmdGET.ExecuteReader
            While r1.Read
                Me.StNum = r1.Item("streetNumber")
                Me.StAddy = r1.Item("StreetName")
                Me.Addy2 = r1.Item("AddressLine2")
                Me.Cty = r1.Item("City")
                Me.St = r1.Item("State")
                Me.Zp = r1.Item("Zip")
                Me.Logo_Dir = r1.Item("LogoDirectory")
                Me.ContactPhNum = r1.Item("ContactPhoneNumber")
                Me.ContactFxNum = r1.Item("ContactFaxNumber")
                Me.CompWebAddy = r1.Item("CompanyWebSite")
                Me.CompName = r1.Item("CompanyName")
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("CompanyInformation", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetInformation")
        End Try
    End Sub
    Public Sub UpdateCompanyInformation()
        Try
            Dim cmdUP As SqlCommand = New SqlCommand("dbo.UpdateCompanyInformation", cnn)
            cmdUP.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@streetNumber", Me.StNumber)
            Dim param2 As SqlParameter = New SqlParameter("@StreetName", Me.StAddress)
            Dim param3 As SqlParameter = New SqlParameter("@AddressLine2", Me.Address_Line_2)
            Dim param4 As SqlParameter = New SqlParameter("@City", Me.City)
            Dim param5 As SqlParameter = New SqlParameter("@State", Me.State)
            Dim param6 As SqlParameter = New SqlParameter("@Zip", Me.Zip)
            Dim param7 As SqlParameter = New SqlParameter("@LogoDirectory", Me.Logo_Directory)
            Dim param8 As SqlParameter = New SqlParameter("@ContactPhoneNumber", Me.ContactPhoneNumber)
            Dim param9 As SqlParameter = New SqlParameter("@ContactFaxNumber", Me.ContactFaxNumber)
            Dim param10 As SqlParameter = New SqlParameter("@CompanyWebsite", Me.Company_WebSite)
            Dim param11 As SqlParameter = New SqlParameter("@CompanyName", Me.Company_Name)
            cmdUP.Parameters.Add(param1)
            cmdUP.Parameters.Add(param2)
            cmdUP.Parameters.Add(param3)
            cmdUP.Parameters.Add(param4)
            cmdUP.Parameters.Add(param5)
            cmdUP.Parameters.Add(param6)
            cmdUP.Parameters.Add(param7)
            cmdUP.Parameters.Add(param8)
            cmdUP.Parameters.Add(param9)
            cmdUP.Parameters.Add(param10)
            cmdUP.Parameters.Add(param11)
            cnn.Open()
            cmdUP.ExecuteNonQuery()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("CompanyInformation", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "UpdateCompanyInformation")
        End Try
    End Sub
End Class
