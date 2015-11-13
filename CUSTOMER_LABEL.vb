Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System


Public Class CUSTOMER_LABEL
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Private C1 As String = ""
    Private C2 As String = ""
    Private StAddy As String = ""
    Private HousePH As String = ""
    Private APhone1 As String = ""
    Private APhone2 As String = ""
    Private AphoneType As String = ""
    Private Aphone2Type As String = ""
    Public CorrectName As String = ""
    Public CorrectStAddress As String = ""
#Region "Customer Properties"
    Public Property Contact1Name() As String
        Get
            Return C1
        End Get
        Set(ByVal value As String)
            C1 = value
        End Set
    End Property
    Public Property Contact2Name() As String
        Get
            Return C2
        End Get
        Set(ByVal value As String)
            C2 = value
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
    Public Property HousePhone() As String
        Get
            Return HousePH
        End Get
        Set(ByVal value As String)
            HousePH = value
        End Set
    End Property
    Public Property AltPhone1() As String
        Get
            Return APhone1
        End Get
        Set(ByVal value As String)
            APhone1 = value
        End Set
    End Property
    Public Property AltPhone2() As String
        Get
            Return APhone2
        End Get
        Set(ByVal value As String)
            APhone2 = value
        End Set
    End Property
    Public Property AltPhone1Type() As String
        Get
            Return AphoneType
        End Get
        Set(ByVal value As String)
            AphoneType = value
        End Set
    End Property
    Public Property AltPhone2Type() As String
        Get
            Return Aphone2Type
        End Get
        Set(ByVal value As String)
            Aphone2Type = value
        End Set
    End Property
#End Region
#Region "Functions"
    Public Function CorrectContactNames(ByVal Contact1FirstName As String, ByVal Contact1LastName As String, ByVal Contact2FirstName As String, ByVal Contact2LastName As String)
        Try
            If Contact2FirstName = "" And Contact2LastName = "" Then
                CorrectName = Contact1FirstName & " " & Contact1LastName
                Return CorrectName
                Exit Function
            End If
            If Contact2LastName = "" And Contact2FirstName <> "" Then

                CorrectName = Contact1FirstName & " and " & Contact2FirstName & " " & Contact1LastName
                Return CorrectName '' comeback
                Exit Function

            End If

            If Contact1LastName = Contact2LastName Then

                CorrectName = Contact1FirstName & " and " & Contact2FirstName & " " & Contact1LastName
                Return CorrectName
                Exit Function
            End If
      

            If Contact1LastName <> Contact2LastName Then
                If Contact2LastName = "" Then
                    CorrectName = Contact1FirstName & " and " & Contact2FirstName & " " & Contact1LastName
                    Return CorrectName
                    Exit Function

                ElseIf Contact2LastName <> "" Then
                    CorrectName = Contact1FirstName & " " & Contact1LastName & " and " & Contact2FirstName & " " & Contact2LastName
                    Return CorrectName
                    Exit Function

                End If
              

            End If
            Return CorrectName
        Catch ex As Exception
            Return CorrectName
            Dim err As New ErrorLogFlatFile
            err.WriteLog("CUSTOMER_LABEL", "ByVal Contact1FirstName As String, ByVal Contact1LastName As String, ByVal Contact2FirstName As String, ByVal Contact2LastName As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CorrectContactNames")
        End Try

    End Function
    Public Function CorrectStreetAddress(ByVal Staddy As String, ByVal City As String, ByVal State As String, ByVal Zip As String)
        CorrectStAddress = Staddy & vbCrLf & City & ", " & State & " " & Zip
        Return CorrectStAddress
    End Function
#End Region
#Region "New Object"
    Public Sub GetINFO(ByVal LeadNumber As String)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("dbo.GetContactLabel", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", LeadNumber)
            cmdGET.Parameters.Add(param1)
            cmdGET.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            While r1.Read
                Me.Contact1Name = CorrectContactNames(r1.Item(0).ToString, r1.Item(1).ToString, r1.Item(2).ToString, r1.Item(3).ToString)
                Me.StAddress = CorrectStreetAddress(r1.Item(4).ToString, r1.Item(5).ToString, r1.Item(6).ToString, r1.Item(7))
                Me.HousePhone = r1.Item(8).ToString
                Me.AltPhone1 = r1.Item(9).ToString
                Me.AltPhone2 = r1.Item(10).ToString
                Me.AltPhone1Type = r1.Item(11).ToString
                Me.AltPhone2Type = r1.Item(12).ToString
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("CUSTOMER_LABEL", "ByVal LeadNumber as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetINFO")
        End Try

    End Sub
#End Region
    

End Class
