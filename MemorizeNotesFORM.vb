Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Public Class MemorizeNotes
 


    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Try

        Catch ex As Exception

        End Try
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.MemorizeAppt", cnn)
        Dim r1 As SqlDataReader
        Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param2 As SqlParameter = New SqlParameter("@Form", STATIC_VARIABLES.ActiveChild.Name)
        Dim param3 As SqlParameter = New SqlParameter("@ID", STATIC_VARIABLES.CurrentID)
        Dim param4 As SqlParameter = New SqlParameter("@Notes", Me.RichTextBox1.Text)
        Dim param5 As SqlParameter = New SqlParameter("@Group", Me.cboGroup.Text)

        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        cmdGet.Parameters.Add(param5)
        cmdGet.CommandType = CommandType.StoredProcedure
        cnn.Open()
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r1.Read()
        r1.Close()
        cnn.Close()
        Me.Close()
        Me.Dispose()
        If STATIC_VARIABLES.ActiveChild.Name = "WCaller" Then
            Dim c As New WarmCalling.MyApptsTab.Populate(WCaller.cboFilter.Text)
        ElseIf STATIC_VARIABLES.ActiveChild.Name = "Sales" Then
            Sales.PopulateMemorized()
        End If

    End Sub

    Private Sub MemorizeNotes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Populate_Groups()
    End Sub
    Private Sub Populate_Groups()
        Me.cboGroup.Items.Clear()
        Me.cboGroup.Items.Add("<Add New>")
        Me.cboGroup.Items.Add("_________________________________________________")
        Me.cboGroup.Items.Add("")

        Try
            Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.PopGroups", cnn)
            Dim r1 As SqlDataReader
            Dim param2 As SqlParameter = New SqlParameter("@Dept", STATIC_VARIABLES.ActiveChild.Name)
            cmdGet.Parameters.Add(param2)
            cmdGet.CommandType = CommandType.StoredProcedure
            cnn.Open()
            r1 = cmdGet.ExecuteReader
            While r1.Read()
                Me.cboGroup.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Add_Group()


        Dim i As String = ""
        i = InputBox$("Enter a new Group Name here.", "Add New Group")

        If i = "" Then
            Me.cboGroup.SelectedItem = Nothing
            Me.cboGroup_SelectedValueChanged(Nothing, Nothing)
            Exit Sub
        End If
        Try
            Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsGroup", cnn)
            Dim cmdget As SqlCommand = New SqlCommand("dbo.PopGroups", cnn)
            Dim param2 As SqlParameter = New SqlParameter("@Dept", STATIC_VARIABLES.ActiveChild.Name)
            Dim param3 As SqlParameter = New SqlParameter("@Dept", STATIC_VARIABLES.ActiveChild.Name)
            Dim param1 As SqlParameter = New SqlParameter("@Group", i)
            cmdINS.Parameters.Add(param1)
            cmdINS.Parameters.Add(param2)
            cmdINS.CommandType = CommandType.StoredProcedure
            cmdget.Parameters.Add(param3)
            cmdget.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            Dim r2 As SqlDataReader
            Me.cboGroup.Items.Clear()
            Me.cboGroup.Items.Add("<Add New>")
            Me.cboGroup.Items.Add("_________________________________________________")
            Me.cboGroup.Items.Add("")
            cnn.Open()
            r2 = cmdINS.ExecuteReader
            r2.Close()

            r1 = cmdget.ExecuteReader
            While r1.Read
                Me.cboGroup.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
            Me.cboGroup.SelectedItem = i
        Catch ex As Exception

        End Try
  


    End Sub

    Private Sub cboGroup_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroup.SelectedValueChanged
        If Me.cboGroup.SelectedItem <> Nothing Then
            Me.lblgroup.Text = Me.cboGroup.Text
            Me.lblgroup.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblgroup.ForeColor = Color.Black
        ElseIf Me.cboGroup.SelectedItem = "" Or Me.cboGroup.SelectedItem = "_________________________________________________" Or Me.cboGroup.SelectedItem = "<Add New>" Then
            Me.lblgroup.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblgroup.ForeColor = Color.Gray
            Me.lblgroup.Text = "Add to Group (Optional)"
        End If



        If Me.cboGroup.SelectedItem = "<Add New>" Then
            Me.cboGroup.SelectedItem = Nothing
            Me.Add_Group()
        ElseIf Me.cboGroup.SelectedItem = "_________________________________________________" Then
            Me.cboGroup.SelectedItem = Nothing

        ElseIf Me.cboGroup.SelectedItem = "" Then
            Me.cboGroup.SelectedItem = Nothing

        End If

    End Sub

    Private Sub lblgroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblgroup.Click
        Me.cboGroup.DroppedDown = True
    End Sub
End Class
