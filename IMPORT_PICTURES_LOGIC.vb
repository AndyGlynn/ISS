Imports System.Data
Imports System.Data.Sql
Imports System
Imports System.Data.SqlClient


Public Class IMPORT_PICTURES_LOGIC
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Private ACR As String = ""

    Public Property ProductAcronym() As String
        Get
            Return ACR
        End Get
        Set(ByVal value As String)
            ACR = value
        End Set
    End Property
    Public Sub GetProducts()
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("dbo.GetProducts", cnn)
            cmdGET.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader(CommandBehavior.CloseConnection)
            ImportPictures.cboProductSel.Items.Clear()
            While r1.Read
                ImportPictures.cboProductSel.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("IMPORT_PICTURE_LOGIC", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetProducts")

        End Try

    End Sub
    Public Sub PullACRO(ByVal Product As String)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("dbo.PullAcro", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Product", Product)
            cmdGET.Parameters.Add(param1)
            cmdGET.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader(CommandBehavior.CloseConnection)
            While r1.Read
                Me.ProductAcronym = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("IMPORT_PICTURE_LOGIC", "ByVal Product as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "PullACRO")

        End Try

    End Sub
End Class
