Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System


Public Class AssignAlert
    Public tt As ToolTip
    Public ep As New ErrorProvider
    Public RemoveErrP As Boolean = False
    Private Sub txtLeadNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLeadNumber.KeyPress
        tt = New ToolTip
        tt.RemoveAll()
        Select Case e.KeyChar
            Case "/"c, "\"c, "!"c, "#"c, "$"c, "%"c, "^"c, "&"c, "("c, ")"c, "?"c, "<"c, ">"c, "@"c
                tt.Show("!@#$%^&*()?<>\/  Are illegal characters to insert.", Me.txtLeadNumber, 3000)
                Me.txtLeadNumber.Text = ""
                Exit Select
        End Select
    End Sub

    Private Sub txtLeadNumber_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadNumber.TextChanged
        Dim str As String = ""
        str = Me.txtLeadNumber.Text
        Dim b As String = Me.txtLeadNumber.Text
        Dim c As Char
        For Each c In b
            Select Case c
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"

                Case Else
                    Me.txtLeadNumber.Text = ""
                    Me.ErrorProvider1.SetError(Me.txtLeadNumber, "This Field Only Accepts Numbers")
                    Exit Sub
            End Select
        Next
        If str.ToString.Length <= 2 Then
            ' Me.lblContactInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            ' Me.lblContactInfo.Text = "No Contact Information"
            Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"

            Exit Sub
        End If
        If str.ToString.ToString <= 0 Then
            'Me.lblContactInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            'Me.lblContactInfo.Text = "No Contact Information"
            Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"

        End If

        If str.ToString.Length > 2 Then
          


            ValidateLeadNumber(str)
            If Me.RemoveErrP = True Then
                Dim a As New CUSTOMER_LABEL
                a.GetINFO(Me.txtLeadNumber.Text)
                Me.lblPhoneInfo.Font = New Font("Tahoma", 12.25!, FontStyle.Bold, GraphicsUnit.Pixel, CType(0, Byte))
                Me.lblPhoneInfo.Text = a.Contact1Name & vbCrLf & a.StAddress & vbCrLf & a.HousePhone & vbCrLf & a.AltPhone1 & "     " & a.AltPhone1Type & vbCrLf & a.AltPhone2 & "     " & a.AltPhone2Type
                Me.ep.Clear()
            End If
        End If
    End Sub
    Private Sub ValidateLeadNumber(ByVal LeadNumber As String)
        ep.BlinkRate = 0
        ep.BlinkStyle = ErrorBlinkStyle.NeverBlink
        ep.Clear()
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdVAL As SqlCommand = New SqlCommand("SELECT COUNT(ID) From iss.dbo.enterlead where ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", LeadNumber)
        cnn.Open()
        cmdVAL.Parameters.Add(param1)
        Dim r1 As SqlDataReader
        r1 = cmdVAL.ExecuteReader
        Try
            While r1.Read
                If r1.Item(0) <= 0 Then
                    'MsgBox("invalid lead num.")
                    ep.SetError(Me.txtLeadNumber, "Invalid Record ID")
                    Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
                    ' Me.lblContactInfo.Text = "No Contact Information"
                    Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"

                    Me.RemoveErrP = False
                ElseIf r1.Item(0) >= 1 Then
                    Me.RemoveErrP = True
                    Me.ep.Clear()
                End If
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception

        End Try

    End Sub
    Private Sub GetAutoCompleteSource()
        'Me.cboUser.AutoCompleteCustomSource.Clear()
        Me.cboUser.Items.Clear()
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGETUSR As SqlCommand = New SqlCommand("SELECT UserFirstName, UserLastName from iss.dbo.userpermissiontable", cnn)
        Dim ArUsers As New ArrayList
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdGETUSR.ExecuteReader
        While r1.Read
            ArUsers.Add(r1.Item(0) & " " & r1.Item(1))
        End While
        r1.Close()
        cnn.Close()
        Dim g As Integer = 0
        For g = 0 To ArUsers.Count - 1
            If ArUsers(g).ToString = "Admin Admin" Then
                Me.cboUser.Items.Add("Admin")
            ElseIf ArUsers(g).ToString <> "Admin" Then
                Me.cboUser.Items.Add(ArUsers(g).ToString)
            End If
        Next
    End Sub

    Private Sub AssignAlert_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.MdiParent = Main
        Me.txtLeadNumber.Text = STATIC_VARIABLES.CurrentID.ToString
        GetAutoCompleteSource()
        If Me.txtLeadNumber.Text = "" Then
            Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"

        End If
        Me.cboUser.Text = STATIC_VARIABLES.CurrentUser


        Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdget As SqlCommand = New SqlCommand("dbo.GetAlertNotes", cnn)
        cmdget.CommandType = CommandType.StoredProcedure
        Dim r1 As SqlDataReader
        Me.cboautonotes.Items.Clear()
        Me.cboautonotes.Items.Add("<Add New>")
        Me.cboautonotes.Items.Add("___________________________________________________")

        cnn.Open()
        r1 = cmdget.ExecuteReader
        While r1.Read
            Me.cboautonotes.Items.Add(r1.Item(0))
        End While
        r1.Close()
        cnn.Close()
    End Sub

   

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.txtLeadNumber.Text = ""
        Me.ep.Clear()
        Me.cboUser.Items.Clear()
        Me.cboUser.Text = ""
        Me.rtfNotes.Text = ""
        Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
        Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cnt As Integer = 0

        Me.ErrorProvider1.Clear()
        If Me.txtLeadNumber.Text = "" Then
            Me.ErrorProvider1.SetError(Me.txtLeadNumber, "Required Field")
            cnt += 1
        End If
        If Me.cboUser.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboUser, "Required Field")
            cnt += 1
        End If
        If Me.txtLeadNumber.Text <> "" And Me.lblPhoneInfo.Text.Contains("No Contact") Then
            Me.ErrorProvider1.SetError(Me.txtLeadNumber, "You Must Supply a Valid Record ID")
            cnt += 1
        End If
        Dim a As String = Me.txtLeadNumber.Text
        Dim c As Char
        For Each c In a
            Select Case c
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"

                Case Else

                    Me.ErrorProvider1.SetError(Me.txtLeadNumber, "This Field Only Accepts Numbers")
                    cnt += 1
            End Select
        Next

        If cnt >= 1 Then
            Exit Sub
        End If




        Dim s = Split(Me.dtpExecutionDate.Value.ToString, " ")
        Dim dt = s(0) & " 12:00:00 AM"
        'Dim s2 = Split(Me.dtpApptTime.Value.ToString, " ")
        Dim tm = "1/1/1900 " & Me.dtpTime.Value.ToShortTimeString
        Dim LN As Integer = CType(Me.txtLeadNumber.Text, Integer)







        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand
        cmdGet = New SqlCommand("dbo.InsertAlert", cnn)
        Dim r1 As SqlDataReader
        Dim param1 As SqlParameter = New SqlParameter("@LeadNum", LN)
        Dim param2 As SqlParameter = New SqlParameter("@UserName", Me.cboUser.Text.ToString)
        Dim param3 As SqlParameter = New SqlParameter("@AlertTime", tm)
        Dim param4 As SqlParameter = New SqlParameter("@ExecutionDate", dt)
        Dim param5 As SqlParameter = New SqlParameter("@Notes", Me.rtfNotes.Text.ToString)
        Dim param6 As SqlParameter = New SqlParameter("@AssignedBy", STATIC_VARIABLES.CurrentUser)
        Dim param7 As SqlParameter = New SqlParameter("@Completed", "False")
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        cmdGet.Parameters.Add(param5)
        cmdGet.Parameters.Add(param6)
        cmdGet.Parameters.Add(param7)
        cnn.Open()
        r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r1.Read()
        r1.Close()
        cnn.Close()
        Me.Close()
        Me.Dispose()

    End Sub
    
    Private Sub cboAutonotes_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAutonotes.SelectedValueChanged
        Dim i As String
        If Me.cboAutonotes.SelectedItem = "<Add New>" Then
            Me.cboAutonotes.Text = Nothing

            i = InputBox$("Enter a new ""Auto Note"" here.", "Save Auto Note")

            If i = "" Then
                MsgBox("You must enter Text to save this Auto Note!", MsgBoxStyle.Exclamation, "No Text Supplied")
                Exit Sub
            End If
            Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsAlertNotes", cnn)
            Dim cmdget As SqlCommand = New SqlCommand("dbo.GetAlertNotes", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Notes", i)
            cmdINS.Parameters.Add(param1)
            cmdINS.CommandType = CommandType.StoredProcedure
            cmdget.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            Dim r2 As SqlDataReader
            Me.cboAutonotes.Items.Clear()
            Me.cboAutonotes.Items.Add("<Add New>")
            Me.cboAutonotes.Items.Add("___________________________________________________")
            cnn.Open()
            r2 = cmdINS.ExecuteReader
            r2.Close()
            r1 = cmdget.ExecuteReader
            While r1.Read
                Me.cboAutonotes.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()

            Me.cboAutonotes.Text = i

        End If
        If Me.cboAutonotes.Text = i Then
            Exit Sub
        End If





        Dim x
        Dim y
        y = Me.rtfNotes.Text.Length
        x = Me.cboAutonotes.Text
        If x = "" Then
            Exit Sub
        End If
        If x = "<Add New>" Then
            Exit Sub
        ElseIf x = "___________________________________________________" Then
            Me.cboAutonotes.Text = ""

            Exit Sub
        End If

        Select Case y
            Case Is < 1
                Me.rtfNotes.Text = x
                Exit Select
            Case Is > 0
                Me.rtfNotes.Text = (Me.rtfNotes.Text & ", " & x)
        End Select
    End Sub

    Private Sub lblAutonotes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblAutonotes.Click
        Me.cboAutonotes.Focus()
        Me.cboAutonotes.DroppedDown = True
    End Sub
End Class
