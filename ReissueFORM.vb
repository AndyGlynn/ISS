Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Public Class Reissue

    Private Sub cboSalesResults_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSalesResults.GotFocus
        Dim x
        x = Me.cboRep1.Text.Length
        Select Case x
            Case Is < 1
                Me.cboSalesResults.Items.Clear()
                Me.cboSalesResults.Items.Add("Not Issued")
                Exit Select
            Case Is > 0
                Me.cboSalesResults.Items.Clear()
                Me.cboSalesResults.Items.Add("Reset")
                Me.cboSalesResults.Items.Add("Not Hit")
        End Select
    End Sub







    Private Sub cboautonotes_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonotes.SelectedValueChanged
        If Me.cboautonotes.Text <> "" Then
            Me.lblautonotes.Text = Me.cboautonotes.Text
            Me.lblautonotes.ForeColor = Color.Black
            Me.lblautonotes.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Else
            Me.lblautonotes.ForeColor = Color.Gray
            Me.lblautonotes.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblautonotes.Text = "select auto note here"
       End If
        If Me.cboautonotes.SelectedItem = "___________________________________________________" Then
            Me.cboautonotes.SelectedItem = Nothing
            Exit Sub
        End If



        Dim i As String
        If Me.cboautonotes.SelectedItem = "<Add New>" Then
            Me.cboautonotes.SelectedItem = Nothing

            i = InputBox$("Enter a new ""Auto Note"" here.", "Save Auto Note")

            If i = "" Then
                MsgBox("You must enter Text to save this Auto Note!", MsgBoxStyle.Exclamation, "No Text Supplied")
                Exit Sub
            End If
            Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.INSSalesManagerAutoNote", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Note", i)
            cmdINS.Parameters.Add(param1)
            cmdINS.CommandType = CommandType.StoredProcedure

            Dim r1 As SqlDataReader

            Me.cboautonotes.Items.Clear()
            Me.cboautonotes.Items.Add("<Add New>")
            Me.cboautonotes.Items.Add("___________________________________________________")
            cnn.Open()


            r1 = cmdINS.ExecuteReader
            While r1.Read
                Me.cboautonotes.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
            Me.cboautonotes.Text = i
        End If
        If Me.cboautonotes.Text = i Then
            Exit Sub
        End If
        Dim x
        Dim y
        y = Me.rtfAutoNote.Text.Length
        x = Me.cboautonotes.Text
        If x = "" Then
            Exit Sub
        End If
        If x = "<Add New>" Then
            Exit Sub
        ElseIf x = "___________________________________________________" Then
            Exit Sub
        End If

        Select Case y
            Case Is < 1
                Me.rtfAutoNote.Text = x
                Exit Select
            Case Is > 0
                Me.rtfAutoNote.Text = (Me.rtfAutoNote.Text & ", " & x)
        End Select
    End Sub
    Private Sub Reissue_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.dtpapptdate.Value = Convert.ToDateTime(Confirming.txtApptDate.Text & " 12:00 AM")
        Me.dtpappttime.Value = Convert.ToDateTime("1/1/1900 " & Confirming.txtApptTime.Text)
        Me.CheckBox1.Enabled = False
        Dim dset_Reps As Data.DataSet = New Data.DataSet("Reps")
        Dim da_Reps As SqlDataAdapter = New SqlDataAdapter

        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim r1 As SqlDataReader
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetSalesReps", cnn)
        Dim cmdget2 As SqlCommand = New SqlCommand("dbo.GETSalesManagerAutoNote", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        Try
            cnn.Open()
            da_Reps.SelectCommand = cmdGet
            da_Reps.Fill(dset_Reps, "SalesReps")
            Dim cnt As Integer = 0
            cnt = dset_Reps.Tables(0).Rows.Count
            Select Case cnt
                Case Is <= 0
                    Me.cboRep1.Items.Clear()
                    Me.cboRep1.Items.Add("")
                    Me.cboRep2.Items.Clear()
                    Me.cboRep2.Items.Add("")

                    Exit Select
                Case Is >= 1
                    Me.cboRep1.Items.Clear()

                    Me.cboRep2.Items.Clear()
                    Me.cboRep2.Items.Add("")
                    Dim b
                    For b = 0 To dset_Reps.Tables(0).Rows.Count - 1
                        Me.cboRep1.Items.Add(dset_Reps.Tables(0).Rows(b).Item(0) & " " & dset_Reps.Tables(0).Rows(b).Item(1))
                        Me.cboRep2.Items.Add(dset_Reps.Tables(0).Rows(b).Item(0) & " " & dset_Reps.Tables(0).Rows(b).Item(1))
                    Next
                    Exit Select
            End Select
            r1 = cmdget2.ExecuteReader
            Me.cboautonotes.Items.Clear()
            Me.cboautonotes.Items.Add("<Add New>")
            Me.cboautonotes.Items.Add("___________________________________________________")

            While r1.Read
                Me.cboautonotes.Items.Add(r1.Item(0))
            End While

            r1.Close()
            cnn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Dim rep As String = Confirming.lvSales.SelectedItems(0).SubItems(3).Text
        Dim rep2 As String = ""
        If rep.Contains(" & ") = True Then
            Dim sp = Split(rep, " & ")
            rep = sp(0)
            rep2 = sp(1)
        End If
        Me.cboRep1.Text = rep
        Me.cboRep2.Text = rep2
        Dim s = Split(Confirming.txtContact1.Text, " ")
        Dim s2 = Split(Confirming.txtContact2.Text, " ")
        If Confirming.txtContact2.Text = " " Then
            Me.cboSpokeWith.Items.Add(s(0))
        ElseIf Confirming.txtContact2.Text <> " " Then
            Me.cboSpokeWith.Items.Add(s(0))
            Me.cboSpokeWith.Items.Add(s2(0))
            Me.cboSpokeWith.Items.Add(s(0) & " & " & s2(0))
        End If
    End Sub
    Private Sub rtfAutoNote_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtfAutoNote.GotFocus
        Me.lblsalesnotes.Visible = False
    End Sub

    Private Sub rtfAutoNote_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtfAutoNote.LostFocus
        Dim x
        x = Me.rtfAutoNote.Text.Length
        Select Case x
            Case Is < 1
                Me.lblsalesnotes.Visible = True
                Exit Select
            Case Is > 0
                Me.lblsalesnotes.Visible = False
        End Select
    End Sub

    Private Sub rtfAutoNote_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtfAutoNote.TextChanged
        Me.lblsalesnotes.Visible = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim day = Me.dtpapptdate.Value.DayOfWeek
        If day = 0 Then
            day = "Sunday"
        ElseIf day = 1 Then
            day = "Monday"
        ElseIf day = 2 Then
            day = "Tuesday"
        ElseIf day = 3 Then
            day = "Wednesday"
        ElseIf day = 4 Then
            day = "Thursday"
        ElseIf day = 5 Then
            day = "Friday"
        ElseIf day = 6 Then
            day = "Saturday"
        End If



        If Me.cboSalesResults.Text = "" And Me.GroupBox1.Enabled = True And Me.cboSpokeWith.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboSalesResults, "Required Field")
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "Required Field")
            Exit Sub
        ElseIf Me.cboSalesResults.Text <> "" And Me.GroupBox1.Enabled = True And Me.cboSpokeWith.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "Required Field")
            Exit Sub
        ElseIf Me.cboSalesResults.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboSalesResults, "Required Field")
            Exit Sub
        End If

        Dim Manager As String
        If Me.cboRep1.Text <> "" Then
            Dim r = Split(Me.cboRep1.Text, " ")
            Manager = Trim(r(1))
        Else
            Manager = ""
        End If


        Dim Description As String
        If Me.dtpapptdate.Value.ToShortDateString = Confirming.dpSales.Value.ToShortDateString Then
            If Me.cboSalesResults.Text = "Not Hit" Then
                If Me.cboRep2.Text = "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & ", which resulted in a ""Not Hit""." & " Logged by " & STATIC_VARIABLES.CurrentUser
                ElseIf Me.cboRep2.Text <> "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & " and " & Me.cboRep2.Text & ", which resulted in a ""Not Hit""." & " Logged by " & STATIC_VARIABLES.CurrentUser
                End If
            ElseIf Me.cboSalesResults.Text = "Reset" Then
                If Me.cboRep2.Text = "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & ", which resulted in a ""Reset""." & " Logged by " & STATIC_VARIABLES.CurrentUser
                ElseIf Me.cboRep2.Text <> "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & " and " & Me.cboRep2.Text & ", which resulted in a ""Reset""." & " Logged by " & STATIC_VARIABLES.CurrentUser
                End If
            ElseIf Me.cboSalesResults.Text = "Not Issued" Then
                Description = "Appt. was not issued. Logged by " & STATIC_VARIABLES.CurrentUser
            End If
        ElseIf Me.dtpapptdate.Value.ToShortDateString <> Confirming.dpSales.Value.ToShortDateString Then
            If Me.cboSalesResults.Text = "Not Hit" Then
                If Me.cboRep2.Text = "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & ", which resulted in a ""Not Hit""." & " Rescheduled by " & STATIC_VARIABLES.CurrentUser & " with " & Me.cboSpokeWith.Text & " for " & day & ", " & Me.dtpapptdate.Value.ToShortDateString & " at " & Me.dtpappttime.Value.ToShortTimeString
                ElseIf Me.cboRep2.Text <> "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & " and " & Me.cboRep2.Text & ", which resulted in a ""Not Hit""." & " Rescheduled by " & STATIC_VARIABLES.CurrentUser & " with " & Me.cboSpokeWith.Text & " for " & day & ", " & Me.dtpapptdate.Value.ToShortDateString & " at " & Me.dtpappttime.Value.ToShortTimeString
                End If
            ElseIf Me.cboSalesResults.Text = "Reset" Then
                If Me.cboRep2.Text = "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & ", which resulted in a ""Reset""." & " Rescheduled by " & STATIC_VARIABLES.CurrentUser & " with " & Me.cboSpokeWith.Text & " for " & day & ", " & Me.dtpapptdate.Value.ToShortDateString & " at " & Me.dtpappttime.Value.ToShortTimeString
                ElseIf Me.cboRep2.Text <> "" Then
                    Description = "Appt. issued to " & Me.cboRep1.Text & " and " & Me.cboRep2.Text & ", which resulted in a ""Reset""." & " Rescheduled by " & STATIC_VARIABLES.CurrentUser & " with " & Me.cboSpokeWith.Text & " for " & day & ", " & Me.dtpapptdate.Value.ToShortDateString & " at " & Me.dtpappttime.Value.ToShortTimeString
                End If
            ElseIf Me.cboSalesResults.Text = "Not Issued" Then
                Description = "Appt. was not issued. Rescheduled by " & STATIC_VARIABLES.CurrentUser & " with " & Me.cboSpokeWith.Text & " for " & day & ", " & Me.dtpapptdate.Value.ToShortDateString & " at " & Me.dtpappttime.Value.ToShortTimeString
            End If
        End If

        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim r1 As SqlDataReader
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.InsSalesResultConfirming", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@Result", Me.cboSalesResults.Text)
        Dim param2 As SqlParameter = New SqlParameter("@Date", Me.dtpapptdate.Value.ToString)
        Dim param3 As SqlParameter = New SqlParameter("@time", Me.dtpappttime.Value.ToString)
        Dim param4 As SqlParameter = New SqlParameter("@user", STATIC_VARIABLES.CurrentUser)
        Dim param5 As SqlParameter = New SqlParameter("@ID", Confirming.lvSales.SelectedItems(0).Text)
        Dim param6 As SqlParameter = New SqlParameter("@Description", Description)
        Dim param7 As SqlParameter = New SqlParameter("@note", Me.rtfAutoNote.Text)
        Dim param8 As SqlParameter = New SqlParameter("@Rep1", Me.cboRep1.Text)
        Dim param9 As SqlParameter = New SqlParameter("@Rep2", Me.cboRep2.Text)
        Dim param10 As SqlParameter = New SqlParameter("@LastName", Manager)
        Dim param11 As SqlParameter = New SqlParameter("@SpokeWith", Me.cboSpokeWith.Text)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        cmdGet.Parameters.Add(param5)
        cmdGet.Parameters.Add(param6)
        cmdGet.Parameters.Add(param7)
        cmdGet.Parameters.Add(param8)
        cmdGet.Parameters.Add(param9)
        cmdGet.Parameters.Add(param10)
        cmdGet.Parameters.Add(param11)
        cnn.Open()
        r1 = cmdGet.ExecuteReader
        r1.Read()
        r1.Close()
        cnn.Close()
        Dim c As New ConfirmingData
        c.Populate("Dispatch", Confirming.cboSalesPLS.Text, Confirming.cboSalesSLS.Text, Confirming.dpSales.Value.ToString, "Populate")
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub lblautonotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblautonotes.Click
        Me.cboautonotes.Focus()
        Me.cboautonotes.DroppedDown = True
    End Sub

    Private Sub lblsalesnotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblsalesnotes.Click
        Me.rtfAutoNote.Focus()
    End Sub


    Private Sub cboSalesResults_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSalesResults.SelectedValueChanged
        Dim x
        x = Me.cboSalesResults.Text.Length
        Select Case x
            Case Is < 1
                Me.GroupBox1.Enabled = False
                Me.CheckBox1.CheckState = False
                Me.CheckBox1.Enabled = False
                Exit Select
            Case Is > 0
                Me.GroupBox1.Enabled = False
                Me.CheckBox1.CheckState = False
                Me.CheckBox1.Enabled = True
        End Select
    End Sub

    Private Sub cboRep1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRep1.SelectedIndexChanged
        If Me.cboRep1.Text <> "" And Me.cboSalesResults.Text = "Not Issued" Then
            Me.cboSalesResults.SelectedItem = Nothing
        ElseIf Me.cboRep1.Text = "" And (Me.cboSalesResults.Text <> "Reset" Or Me.cboSalesResults.Text <> "Not Hit") Then
            Me.cboSalesResults.SelectedItem = Nothing
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Me.GroupBox1.Enabled = True
        Else
            Me.GroupBox1.Enabled = False
        End If
    End Sub

    Private Sub cboautonotes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboautonotes.SelectedIndexChanged

    End Sub
End Class
