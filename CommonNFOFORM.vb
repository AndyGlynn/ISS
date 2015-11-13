Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class CommonNFO
    Public RecordID As String = ""
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.lblAddress.Text = ""
        Me.lblID.Text = ""
        Me.lblContact.Text = ""
        Me.lblHousePhone.Text = ""
        Me.lblAltPhone.Text = ""
        Me.lblAltPhone2.Text = ""
        Me.lblPhone1Type.Text = ""
        Me.lblPhone2Type.Text = ""
        Me.Close()
    End Sub

    Private Sub CommonNFO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetINFO(RecordID)
    End Sub
    Private Sub GetINFO(ByVal ID As String)
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT Contact1FirstName,Contact1LastName,Contact2FirstName,Contact2LastName,Staddress,city,state,zip,housephone,altphone1,altphone2,phone1type,phone2type from iss.dbo.enterlead where ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdGET.Parameters.Add(param1)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader
        While r1.Read
            Me.lblAddress.Text = r1.Item("Staddress") & vbCrLf & r1.Item("City") & ", " & r1.Item("State") & ", " & r1.Item("Zip")
            Me.lblContact.Text = r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName") & vbCrLf & r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName")
            Me.lblID.Text = ID
            Me.lblHousePhone.Text = r1.Item("HousePhone")
            Me.lblAltPhone.Text = r1.Item("AltPhone1")
            Me.lblPhone1Type.Text = r1.Item("Phone1Type")
            Me.lblPhone2Type.Text = r1.Item("Phone2Type")
            Me.lblAltPhone2.Text = r1.Item("AltPhone2")
        End While
        r1.Close()
        cnn.Close()

    End Sub

End Class
