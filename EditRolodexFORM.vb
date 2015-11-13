Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class EditRolodex
    Public Target_ID As String = "0"
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Private Sub Edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT * from iss.dbo.companyrolodex where id = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", Target_ID)
        cmdGET.Parameters.Add(param1)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader
        While r1.Read
            Me.txtFirstName.Text = r1.Item("EmpFirstName")
            Me.txtLastName.Text = r1.Item("EmpLastName")
            Me.txtPhoneNumber.Text = r1.Item("PrimaryPhone")
            Me.cboDepartment.SelectedItem = r1.Item("Department")
        End While
        r1.Close()
        cnn.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAction.Click
        '' update the selected ID
        '' then clear out the form
        '' then close the form 
        Select Case Me.btnAction.Text
            Case Is = "Add"
                Dim cmdFIND As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.companyrolodex where EmpFirstName = @EMP and EmpLastName = @EMPL", cnn)
                Dim param44 As SqlParameter = New SqlParameter("@EMP", Me.txtFirstName.Text)
                Dim param55 As SqlParameter = New SqlParameter("@EMPL", Me.txtLastName.Text)
                cmdFIND.Parameters.Add(param44)
                cmdFIND.Parameters.Add(param55)
                Dim r1 As SqlDataReader
                cnn.Open()
                r1 = cmdFIND.ExecuteReader
                Dim cnt As Integer = 0
                While r1.Read
                    cnt = r1.Item(0)
                End While
                cnn.Close()
                If cnt >= 1 Then
                    MsgBox("Duplicate employee exists. Please enter a new employee.", MsgBoxStyle.Critical, "Duplicate Employee")
                    Exit Sub
                ElseIf cnt <= 0 Then
                    Dim cmdINS As SqlCommand = New SqlCommand("INSERT iss.dbo.companyrolodex (EmpFirstName,EmpLastName,PrimaryPhone,Department) " _
                    & " VALUES(@EMPF,@EMPL,@PH,@DEP)", cnn)
                    Dim param60 As SqlParameter = New SqlParameter("@EMPF", Me.txtFirstName.Text)
                    Dim param61 As SqlParameter = New SqlParameter("@EMPL", Me.txtLastName.Text)
                    Dim param62 As SqlParameter = New SqlParameter("@DEP", Me.cboDepartment.Text)
                    Dim param63 As SqlParameter = New SqlParameter("@PH", Me.txtPhoneNumber.Text)
                    cmdINS.Parameters.Add(param60)
                    cmdINS.Parameters.Add(param61)
                    cmdINS.Parameters.Add(param62)
                    cmdINS.Parameters.Add(param63)
                    cnn.Open()
                    cmdINS.ExecuteReader()
                    cnn.Close()
                End If
                Me.Target_ID = "0"
                Me.txtFirstName.Text = ""
                Me.txtLastName.Text = ""
                Me.cboDepartment.Text = ""
                Me.txtPhoneNumber.Text = ""
                Me.Close()
                frmRolodex.ListView1.Items.Clear()
                frmRolodex.PopulateDefaultList(frmRolodex.cboDepartment.Text)
                Exit Select
            Case Is = "Edit"
                Dim cmdUP As SqlCommand = New SqlCommand("UPDATE iss.dbo.companyrolodex " _
                & " SET EmpFirstName = @EMPF, " _
                & "     EmpLastName = @EMPL, " _
                & "     PrimaryPhone = @PPH, " _
                & "     Department = @DEP  WHERE ID = @ID", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@EMPF", Me.txtFirstName.Text)
                Dim param2 As SqlParameter = New SqlParameter("@EMPL", Me.txtLastName.Text)
                Dim param3 As SqlParameter = New SqlParameter("@PPH", Me.txtPhoneNumber.Text)
                Dim param4 As SqlParameter = New SqlParameter("@DEP", Me.cboDepartment.Text)
                Dim param5 As SqlParameter = New SqlParameter("@ID", Me.Target_ID)
                cmdUP.Parameters.Add(param1)
                cmdUP.Parameters.Add(param2)
                cmdUP.Parameters.Add(param3)
                cmdUP.Parameters.Add(param4)
                cmdUP.Parameters.Add(param5)
                cnn.Open()
                cmdUP.ExecuteReader()
                cnn.Close()
                Me.txtFirstName.Text = ""
                Me.txtLastName.Text = ""
                Me.txtPhoneNumber.Text = ""
                Me.cboDepartment.Text = ""
                Me.Target_ID = "0"
                Me.Close()
                frmRolodex.ListView1.Items.Clear()
                frmRolodex.PopulateDefaultList(frmRolodex.cboDepartment.Text)
                Exit Select
        End Select
        Me.Target_ID = "0"
        Me.txtFirstName.Text = ""
        Me.txtLastName.Text = ""
        Me.cboDepartment.Text = ""
        Me.txtPhoneNumber.Text = ""
        Me.btnAction.Text = "Edit"
        Me.Close()
        frmRolodex.PopulateDefaultList(frmRolodex.cboDepartment.Text)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Target_ID = "0"
        Me.txtFirstName.Text = ""
        Me.txtLastName.Text = ""
        Me.cboDepartment.Text = ""
        Me.txtPhoneNumber.Text = ""

        Me.Close()
    End Sub
End Class
