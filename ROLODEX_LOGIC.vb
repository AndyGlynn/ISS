Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System


Public Class ROLODEX_LOGIC
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
#Region "Private Classes"
    Public Class GetEmployeeByDepartment
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub GetEMPLOYEES(ByVal department As String)
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetEmployeeByDepartment", cnn)
                cmdGet.CommandType = CommandType.StoredProcedure
                Dim param1 As SqlParameter = New SqlParameter("@Department", department)
                cmdGet.Parameters.Add(param1)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                Employee_Contacts.lstEmployees.Items.Clear()
                While r1.Read
                    Dim lv As New ListViewItem
                    lv.Text = r1.Item(1) & ", " & r1.Item(0)
                    lv.SubItems.Add(r1.Item(2))
                    Employee_Contacts.lstEmployees.Items.Add(lv)
                End While
                r1.Close()
                cnn.Close()
            Catch ex As Exception
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ROLODEX_LOGIC.GetEmployeeByDepartment", "ByVal department as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetEMPLOYEES")
            End Try
        End Sub
    End Class
    Public Class EditEmployee
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub EditEmployees(ByVal RecID As String, ByVal NewFname As String, ByVal NewLName As String, ByVal PrimaryPhone As String)
            Try
                Dim cmdUP As SqlCommand = New SqlCommand("dbo.EditEmployee", cnn)
                Dim param4 As SqlParameter = New SqlParameter("@ID", RecID)
                Dim param1 As SqlParameter = New SqlParameter("@NewFName", NewFname)
                Dim param2 As SqlParameter = New SqlParameter("@NewLName", NewLName)
                Dim param3 As SqlParameter = New SqlParameter("@PrimaryPhone", PrimaryPhone)
                cmdUP.Parameters.Add(param1)
                cmdUP.Parameters.Add(param2)
                cmdUP.Parameters.Add(param3)
                cmdUP.Parameters.Add(param4)
                cmdUP.CommandType = CommandType.StoredProcedure
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdUP.ExecuteReader(CommandBehavior.CloseConnection)
                r1.Close()
                cnn.Close()

                Dim z As String = Employee_Contacts.cboDepartment.Text
                If z.ToString.Length < 2 Then
                    Employee_Contacts.lstEmployees.Items.Clear()
                    Exit Sub
                End If
                Dim g As New GetEmployeeByDepartment
                g.GetEMPLOYEES(z)
            Catch ex As Exception
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ROLODEX_LOGIC.EditEmployee", "ByVal RecID As String, ByVal NewFname As String, ByVal NewLName As String, ByVal PrimaryPhone As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "EditEmployees")
            End Try
        End Sub
    End Class
    Public Class GetRecIDForEmployee
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public RecID As String = ""
        Public Sub GetRecID(ByVal EmpFname As String, ByVal EmpLName As String)
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetEmpRecID", cnn)
                cmdGet.CommandType = CommandType.StoredProcedure
                Dim param1 As SqlParameter = New SqlParameter("@EmpFName", EmpFname)
                Dim param2 As SqlParameter = New SqlParameter("@EmpLName", EmpLName)
                cnn.Open()
                cmdGet.Parameters.Add(param1)
                cmdGet.Parameters.Add(param2)
                Dim r1 As SqlDataReader
                r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                While r1.Read
                    RecID = r1.Item(0)
                End While
                r1.Close()
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ROLODEX_LOGIC.GetRecIDForEmployee", "ByVal EmpFname As String, ByVal EmpLName As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetRecID")
            End Try
        End Sub
    End Class
    Public Class DeleteEmployee
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub DelEmployee(ByVal RecID As String)
            Try
                Dim cmdDel As SqlCommand = New SqlCommand("Delete iss.dbo.CompanyRolodex WHERE ID = @RECID", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@RecID", RecID)
                cmdDel.Parameters.Add(param1)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdDel.ExecuteReader(CommandBehavior.CloseConnection)
                r1.Close()
                cnn.Close()
                Dim g As New GetEmployeeByDepartment
                g.GetEMPLOYEES(Employee_Contacts.cboDepartment.Text)
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ROLODEX_LOGIC.DeleteEmployee", "ByVal RecID As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "DelEmployee")
            End Try
        End Sub
    End Class
    Public Class InsetEmployee
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub InsertNewEmployee(ByVal FName As String, ByVal LName As String, ByVal Department As String, ByVal PrimaryPhone As String)
            Try
                Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.CompanyRolodex WHERE EmpFirstName = @EMPF and EmpLastName = @EMPL and Department = @DEP", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@EMPF", FName)
                Dim param2 As SqlParameter = New SqlParameter("@EMPL", LName)
                Dim param3 As SqlParameter = New SqlParameter("@DEP", Department)
                cmdCNT.Parameters.Add(param1)
                cmdCNT.Parameters.Add(param2)
                cmdCNT.Parameters.Add(param3)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdCNT.ExecuteReader(CommandBehavior.CloseConnection)
                Dim cnt As Integer = 0
                While r1.Read
                    cnt = r1.Item(0)
                End While
                r1.Close()
                cnn.Close()
                Select Case cnt
                    Case Is >= 1
                        MsgBox("A duplicate employee for '" & LName & ", " & FName & " exists already. Please enter a new employee.", MsgBoxStyle.Exclamation, "ERROR")
                        Exit Sub
                    Case Is <= 0
                        Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertRolodex", cnn)
                        cmdINS.CommandType = CommandType.StoredProcedure
                        Dim param4 As SqlParameter = New SqlParameter("@EmpFirstName", FName)
                        Dim param5 As SqlParameter = New SqlParameter("@EmpLastName", LName)
                        Dim param6 As SqlParameter = New SqlParameter("@Department", Department)
                        Dim param7 As SqlParameter = New SqlParameter("@PrimaryPhone", PrimaryPhone)
                        cmdINS.Parameters.Add(param4)
                        cmdINS.Parameters.Add(param5)
                        cmdINS.Parameters.Add(param6)
                        cmdINS.Parameters.Add(param7)
                        cnn.Open()
                        Dim r2 As SqlDataReader
                        r2 = cmdINS.ExecuteReader(CommandBehavior.CloseConnection)
                        r2.Close()
                        cnn.Close()
                        Dim g As New ROLODEX_LOGIC.GetEmployeeByDepartment
                        g.GetEMPLOYEES(Employee_Contacts.cboDepartment.Text)
                        Exit Select
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ROLODEX_LOGIC.InsetEmployee", "ByVal FName As String, ByVal LName As String, ByVal Department As String, ByVal PrimaryPhone As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertNewEmployee")
            End Try
        End Sub
    End Class
#End Region
End Class
