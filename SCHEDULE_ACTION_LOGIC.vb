
Imports System.Data
Imports System
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class SCHEDULE_ACTION_LOGIC
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Public Sub GetActionList(ByVal Department As String)
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetActions", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Department", Department)
            cmdGet.Parameters.Add(param1)
            cmdGet.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            ScheduleAction.CboScheduledAction.Items.Clear()
            ScheduleAction.CboScheduledAction.Items.Add("<Add New>")
            ScheduleAction.CboScheduledAction.Items.Add("________________________")
            ScheduleAction.CboScheduledAction.Items.Add("")
            While r1.Read
                ScheduleAction.CboScheduledAction.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("SCHEDULE_ACTION_LOGIC.GetActionList", "ByVal Department as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetActionList")

        End Try

    End Sub
    Public Sub InsertNewAction(ByVal Department As String, ByVal Action As String)
        Try
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertAction", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Department", Department)
            Dim param2 As SqlParameter = New SqlParameter("@Action", Action)
            cmdINS.Parameters.Add(param1)
            cmdINS.Parameters.Add(param2)
            cnn.Open()
            cmdINS.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            r1 = cmdINS.ExecuteReader(CommandBehavior.CloseConnection)
            r1.Close()
            cnn.Close()
            GetActionList(Department)
            'ScheduleAction.CboScheduledAction.SelectedItem = Action
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("SCHEDULE_ACTION_LOGIC.GetActionList", "ByVal Department as string, ByVal Action As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertNewAction")

        End Try

    End Sub
#Region "Private Classes"
    Public Class InsertSA
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub InsertNewSchedAction(ByVal LeadNum As String, ByVal Department As String, _
         ByVal AssignedTo As String, ByVal ExecDate As Date, ByVal Notes As String, ByVal AttachedFiles As Boolean, ByVal scheduledAction As String, ByVal Hash As String, ByVal Completed As Boolean)
            Try
                Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.ScheduledTasks " _
                & " WHERE LeadNum = @LEADNUM and Department = @DEP and AssignedTo = @AT and ExecutionDate = @ED and SchedAction = @SA", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@LeadNum", LeadNum)
                Dim param2 As SqlParameter = New SqlParameter("@DEP", Department)
                Dim param3 As SqlParameter = New SqlParameter("@AT", AssignedTo)
                Dim param4 As SqlParameter = New SqlParameter("@ED", ExecDate)
                Dim param5 As SqlParameter = New SqlParameter("@SA", scheduledAction)

                cmdCNT.Parameters.Add(param1)
                cmdCNT.Parameters.Add(param2)
                cmdCNT.Parameters.Add(param3)
                cmdCNT.Parameters.Add(param4)
                cmdCNT.Parameters.Add(param5)
                cnn.Open()
                cmdCNT.CommandType = CommandType.Text
                Dim r1 As SqlDataReader
                Dim cnt As Integer = 0
                r1 = cmdCNT.ExecuteReader(CommandBehavior.CloseConnection)
                While r1.Read
                    cnt = r1.Item(0)
                End While
                r1.Close()
                cnn.Close()
                Select Case cnt
                    Case Is <= 0
                        Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertSA", cnn)
                        Dim param6 As SqlParameter = New SqlParameter("@LeadNum", LeadNum)
                        Dim param7 As SqlParameter = New SqlParameter("@Department", Department)
                        Dim param8 As SqlParameter = New SqlParameter("@AssignedTo", AssignedTo)
                        Dim param9 As SqlParameter = New SqlParameter("@ExecutionDate", ExecDate)
                        Dim param10 As SqlParameter = New SqlParameter("@Notes", Notes)
                        Dim param11 As SqlParameter = New SqlParameter("@AttachedFiles", AttachedFiles)
                        Dim param12 As SqlParameter = New SqlParameter("@SchedAction", scheduledAction)
                        Dim param13 As SqlParameter = New SqlParameter("@Completed", Completed)
                        Dim param15 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
                        '' hash will not be blank as dictated by the logic on the front
                        '' end so there is no need to compensate for a blank value here.
                        Dim param14 As SqlParameter = New SqlParameter("@Hash", Hash)
                        cmdINS.CommandType = CommandType.StoredProcedure
                        cmdINS.Parameters.Add(param6)
                        cmdINS.Parameters.Add(param7)
                        cmdINS.Parameters.Add(param8)
                        cmdINS.Parameters.Add(param9)
                        cmdINS.Parameters.Add(param10)
                        cmdINS.Parameters.Add(param11)
                        cmdINS.Parameters.Add(param12)
                        cmdINS.Parameters.Add(param13)
                        cmdINS.Parameters.Add(param14)
                        cmdINS.Parameters.Add(param15)
                        cnn.Open()
                        Dim r2 As SqlDataReader
                        r2 = cmdINS.ExecuteReader(CommandBehavior.CloseConnection)
                        'ScheduleAction.SAID = r2.Item(0).ToString()
                        r2.Close()
                        cnn.Close()
                        Dim cmdUP As SqlCommand = New SqlCommand("Select max(id) from ScheduledTasks ", cnn)
                        cmdUP.CommandType = CommandType.Text
                        cnn.Open()
                        Dim r3 As SqlDataReader
                        r3 = cmdUP.ExecuteReader(CommandBehavior.CloseConnection)
                        While r3.Read
                            ScheduleAction.SAID = r3.Item(0).ToString
                        End While
                        r3.Close()
                        cnn.Close()

                        Me.Refresh_CustomerHistory()
                        Exit Select
                    Case Is >= 1
                        MsgBox("Duplicate scheduled action exists. Please enter new scheduled task.", MsgBoxStyle.Critical, "ERROR")
                        Exit Sub
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("SCHEDULE_ACTION_LOGIC.InsertSA", "ByVal LeadNum As String, ByVal Department As String, ByVal AssignedTo As String, ByVal ExecDate As Date, ByVal Notes As String, ByVal AttachedFiles As Boolean, ByVal scheduledAction As String, ByVal Hash As String, ByVal Completed As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertNewSchedAction")

            End Try

        End Sub
        Public Sub UpdateSchedAction(ByVal id As String, ByVal LeadNum As String, ByVal Department As String, _
   ByVal AssignedTo As String, ByVal ExecDate As Date, ByVal Notes As String, ByVal AttachedFiles As Boolean, ByVal scheduledAction As String, ByVal Hash As String)
            ScheduleAction.SAID = id
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.UpdateSA", cnn)
            Dim param6 As SqlParameter = New SqlParameter("@LeadNum", LeadNum)
            Dim param7 As SqlParameter = New SqlParameter("@Department", Department)
            Dim param8 As SqlParameter = New SqlParameter("@AssignedTo", AssignedTo)
            Dim param9 As SqlParameter = New SqlParameter("@ExecutionDate", ExecDate)
            Dim param10 As SqlParameter = New SqlParameter("@Notes", Notes)
            Dim param11 As SqlParameter = New SqlParameter("@AttachedFiles", AttachedFiles)
            Dim param12 As SqlParameter = New SqlParameter("@SchedAction", scheduledAction)
            Dim param13 As SqlParameter = New SqlParameter("@ID", id)
            Dim param15 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
            '' hash will not be blank as dictated by the logic on the front
            '' end so there is no need to compensate for a blank value here.
            Dim param14 As SqlParameter = New SqlParameter("@Hash", Hash)
            cmdINS.CommandType = CommandType.StoredProcedure
            cmdINS.Parameters.Add(param6)
            cmdINS.Parameters.Add(param7)
            cmdINS.Parameters.Add(param8)
            cmdINS.Parameters.Add(param9)
            cmdINS.Parameters.Add(param10)
            cmdINS.Parameters.Add(param11)
            cmdINS.Parameters.Add(param12)
            cmdINS.Parameters.Add(param13)
            cmdINS.Parameters.Add(param14)
            cmdINS.Parameters.Add(param15)
            cnn.Open()
            Dim r2 As SqlDataReader
            r2 = cmdINS.ExecuteReader(CommandBehavior.CloseConnection)
            'ScheduleAction.SAID = r2.Item(0).ToString()
            r2.Close()
            cnn.Close()
          

            'Me.Refresh_CustomerHistory()
        End Sub
        Public Sub Update_AF_path(ByVal hash As String, ByVal SAID As String)
            Try
                Dim cmdUP As SqlCommand = New SqlCommand("update ScheduledTasks set AttachedHashValue = '" & Hash & "' where id = '" & SAID & "'", cnn)
                cmdUP.CommandType = CommandType.Text
                cnn.Open()
                Dim r2 As SqlDataReader
                r2 = cmdUP.ExecuteReader(CommandBehavior.CloseConnection)
                r2.Close()
                cnn.Close()
            Catch ex As Exception

            End Try
        End Sub
        Private Sub Refresh_CustomerHistory()
            Dim x As New CustomerHistory
            Dim tscbo As ToolStripComboBox = Nothing
            Select Case STATIC_VARIABLES.ActiveChild.Name
                Case "Confirming"
                    tscbo = Confirming.TScboCustomerHistory
                Case "Sales"
                    tscbo = Sales.TScboCustomerHistory
                Case "Administration"
                    tscbo = Administration.TScboCustomerHistory
                Case "Finance"
                    tscbo = Finance.TScboCustomerHistory
                Case "WCaller"
                    tscbo = WCaller.TScboCustomerHistory
                Case "Recovery"
                    tscbo = Recovery.TScboCustomerHistory
                Case "PreviousCustomer"
                    tscbo = PreviousCustomer.TScboCustomerHistory
                Case "Installation"
                    tscbo = Installation.TScboCustomerHistory
                Case "ConfirmingSingleRecord"
                    tscbo = ConfirmingSingleRecord.TScboCustomerHistory
                Case "SecondSource"
                    tscbo = SecondSource.TScboCustomerHistory
                Case "MarketingManager"
                    tscbo = MarketingManager.TScboCustomerHistory
                Case "ColdCalling"
                    tscbo = ColdCalling.TScboCustomerHistory
            End Select
            If tscbo Is Nothing Then
                Exit Sub
            End If
            x.SetUp(STATIC_VARIABLES.ActiveChild, STATIC_VARIABLES.CurrentID, tscbo)
        End Sub
    End Class

#End Region
End Class
