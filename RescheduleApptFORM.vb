Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class RescheduleAppt
    Public frm As String
    Public ID As String

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim s = Split(Me.dtpApptDate.Value.ToString, " ")
        Dim dt = s(0) & " 12:00:00 AM"
        'Dim s2 = Split(Me.dtpApptTime.Value.ToString, " ")
        Dim tm = "1/1/1900 " & Me.dtpApptTime.Value.ToShortTimeString





        If Me.dtpApptDate.Value < Date.Today Then
            MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Move Appt.")
            Exit Sub
        End If
        Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdINS As SqlCommand = New SqlCommand("dbo.MoveAppt", cnn)

        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        Dim param2 As SqlParameter = New SqlParameter("@Date", dt)
        Dim param3 As SqlParameter = New SqlParameter("@Time", tm)
        Dim param4 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        cmdINS.Parameters.Add(param1)
        cmdINS.Parameters.Add(param2)
        cmdINS.Parameters.Add(param3)
        cmdINS.Parameters.Add(param4)
        cmdINS.CommandType = CommandType.StoredProcedure
        Dim r1 As SqlDataReader
        cnn.Open()
        r1 = cmdINS.ExecuteReader
        r1.Read()
        r1.Close()
        cnn.Close()
        If frm = "Confirming" Then
            Dim c As New ConfirmingData
            If Confirming.Tab = "Confirm" Then

                c.Populate(Confirming.Tab, Confirming.cboConfirmingPLS.Text, Confirming.cboConfirmingSLS.Text, Confirming.dpConfirming.Value.ToString, "Populate")
            Else
                c.Populate(Confirming.Tab, Confirming.cboSalesPLS.Text, Confirming.cboSalesSLS.Text, Confirming.dpSales.Value.ToString, "Populate")
            End If

        End If









            Me.Close()
            Me.Dispose()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

   

    Private Sub RescheduleAppt_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
