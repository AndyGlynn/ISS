Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Public Class CNGApptTime

    Private Sub CNGApptTime_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtApptTime.Text = Confirming.txtApptTime.Text
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.dtpApptTime.Value.ToShortTimeString = Me.txtApptTime.Text Then
            Me.Close()
            Exit Sub
        End If
        Dim Description As String = "At " & DateTime.Now.ToShortTimeString & " " & STATIC_VARIABLES.CurrentUser & " changed Appt. Time from " & Me.txtApptTime.Text & " to " & Me.dtpApptTime.Value.ToShortTimeString



        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdChange As SqlCommand
        cmdChange = New SqlCommand("dbo.CNGApptTime", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", Confirming.lvSales.SelectedItems(0).Text)
        Dim param2 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param3 As SqlParameter = New SqlParameter("@Description", Description)
        Dim param4 As SqlParameter = New SqlParameter("@ApptTime", Me.dtpApptTime.Value)

        cmdChange.CommandType = CommandType.StoredProcedure
        cmdChange.Parameters.Add(param1)
        cmdChange.Parameters.Add(param2)
        cmdChange.Parameters.Add(param3)
        cmdChange.Parameters.Add(param4)
        cnn.Open()
        Dim R1 = cmdChange.ExecuteReader
        R1.Read()
        R1.Close()
        cnn.Close()
        Dim c As New ConfirmingData
        c.Populate("Dispatch", Confirming.cboSalesPLS.Text, Confirming.cboSalesSLS.Text, Confirming.dpSales.Value.ToString, "Refresh")
        Dim c2 As New CustomerHistory
        c2.SetUp(Confirming, Confirming.lvSales.SelectedItems(0).Text, Confirming.TScboCustomerHistory)
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class
