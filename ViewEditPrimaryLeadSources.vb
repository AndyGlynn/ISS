Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System


Public Class ViewEditPrimaryLeadSources

    '' sql connection string
    '' 

    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)

    '' array list
    '' 
    Public ArPLS As New ArrayList
    Public Sub PopulateList()
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("SELECT PrimaryLead from iss.dbo.primaryleadsource", cnn)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            While r1.Read
                ArPLS.Add(r1.Item("PrimaryLead").ToString)
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ViewEditPrimaryLeadSources", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "PopulateList")

        End Try
    End Sub
    Public Sub Add_PLS(ByVal PLS As String)
        Try
            Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.primaryleadsource where PrimaryLead = @PLS", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)
            cmdCNT.Parameters.Add(param1)
            cnn.Open()
            Dim cnt As Integer = 0
            cnt = cmdCNT.ExecuteScalar
            cnn.Close()
            Select Case cnt
                Case Is <= 0
                    '' doesnt exist
                    '' add it
                    '' refresh
                    Try
                        Dim cmdINS As SqlCommand = New SqlCommand("INSERT iss.dbo.primaryleadsource (PrimaryLead) values(@PLS)", cnn)
                        Dim param2 As SqlParameter = New SqlParameter("@PLS", PLS)
                        cmdINS.Parameters.Add(param2)
                        cnn.Open()
                        cmdINS.ExecuteNonQuery()
                        cnn.Close()
                        Exit Select
                    Catch ex As Exception
                        cnn.Close()
                    End Try

                Case Is >= 1
                    '' does exist
                    '' flag it / warn user
                    '' exit sub 
                    MsgBox("Primary Lead Source '" & PLS.ToString & "' already exists. Please insert a new Primary Lead Source.", MsgBoxStyle.Exclamation, "Error Adding Primary Lead Source")
                    Exit Select
            End Select
        Catch ex As Exception
            cnn.Close()
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ViewEditPrimaryLeadSources", "ByVal PLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "PopulateList")

        End Try
    End Sub

End Class
