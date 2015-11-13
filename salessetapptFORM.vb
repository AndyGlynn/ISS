Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient


Public Class salessetappt
    Public Contact1 As String
    Public Contact2 As String
    Public ID As String
    Public OrigApptDate As String
    Public OrigApptTime As String
    Public frm As Form
    Public Confirmed As Boolean = False

  
    Private Sub salessetappt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ShowInTaskbar = False
        Me.AcceptButton = Me.btnsave
        Me.cboSpokeWith.Items.Clear()
        Me.cboSpokeWith.Items.Add(Contact1)
        If Contact2 <> "" Then
            Me.cboSpokeWith.Items.Add(Contact2)
            Me.cboSpokeWith.Items.Add(Contact1 & " & " & Contact2)
        Else
            Me.cboSpokeWith.Text = Me.Contact1
        End If

        Me.btnsave.Focus()
        Dim DD As String = OrigApptDate
        Dim DT As String = OrigApptTime
        DT = "1/1/1900 " & DT
        Convert.ToDateTime(DD)
        Convert.ToDateTime(DT)
        Me.dtpApptDate.Value = DD
        Me.dtpApptTime.Value = DT
        Me.txtDate.Text = ""
        Me.txtDate.Visible = True

        Me.txtTime.Text = ""
        Me.txtTime.Visible = True
        Me.cboSpokeWith.Text = ""
        Me.cboautonotes.SelectedItem = Nothing


        Me.RichTextBox1.Text = ""

        Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdget As SqlCommand = New SqlCommand("dbo.GetSetApptNotes", cnn)
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
    Private Sub cboautonotes_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonotes.SelectedValueChanged
        Dim i As String
        If Me.cboautonotes.SelectedItem = "<Add New>" Then
            Me.cboautonotes.SelectedItem = Nothing

            i = InputBox$("Enter a new ""Auto Note"" here.", "Save Auto Note")

            If i = "" Then
                MsgBox("You must enter Text to save this Auto Note!", MsgBoxStyle.Exclamation, "No Text Supplied")
                Exit Sub
            End If
            Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsSetApptNotes", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Notes", i)
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
        y = Me.RichTextBox1.Text.Length
        x = Me.cboautonotes.Text
        If x = "" Then
            Exit Sub
        End If
        If x = "<Add New>" Then
            Me.lblautonotes.ForeColor = Color.Gray
            Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblautonotes.Text = "Select Auto Note Here"
            'Me.cboautonotes.SelectedItem = Nothing
            Exit Sub
        ElseIf x = "___________________________________________________" Then
            Me.lblautonotes.ForeColor = Color.Gray
            Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblautonotes.Text = "Select Auto Note Here"
            Me.cboautonotes.SelectedItem = Nothing
            Exit Sub
        End If

        Select Case y
            Case Is < 1
                Me.RichTextBox1.Text = x
                Me.lblautonotes.Text = Me.cboautonotes.Text
                Me.lblautonotes.ForeColor = Color.Black
                Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))


                Exit Select
            Case Is > 0
                Me.RichTextBox1.Text = (Me.RichTextBox1.Text & ", " & x)
                Me.lblautonotes.Text = Me.cboautonotes.Text
                Me.lblautonotes.ForeColor = Color.Black
                Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        End Select
        If Me.RichTextBox1.Text <> "" Then
            Me.Label1.Visible = False
        End If

    End Sub


    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        If Me.dtpApptDate.Value < Today And Me.txtDate.Visible = False Then

            Dim er As New ErrorProvider
            er.BlinkStyle = ErrorBlinkStyle.NeverBlink
            er.SetError(Me.dtpApptDate, "Appointment Date Must be for Today or Future!")
            'MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Set Appt.")
            Exit Sub
        End If

        Me.ErrorProvider1.Clear()
        Me.ErrorProvider1.BlinkStyle = ErrorBlinkStyle.NeverBlink
        If Me.cboSpokeWith.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "You Must supply a Name!")
        End If

        'If Me.RichTextBox1.Text = "" Then
        'Me.ErrorProvider1.SetError(Me.RichTextBox1, "You must supply a reason for cancellation!")
        'End If
        If Me.txtDate.Text = "" Then
            Me.ErrorProvider1.SetError(Me.dtpApptDate, "You must supply a new Appt. Date!")
        End If
        If Me.txtTime.Text = "" Then
            Me.ErrorProvider1.SetError(Me.dtpApptTime, "You must supply a new Appt. Time!")
        End If
        Dim Description As String
        If Me.cboSpokeWith.Text <> "" And Me.txtDate.Text <> "" And Me.txtTime.Text <> "" Then
            Dim day = Me.dtpApptDate.Value.DayOfWeek
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






            If Me.txtDate.Text = "" Then

                Me.txtDate.Text = OrigApptDate
            End If
            If Me.txtTime.Text = "" Then
                Me.txtTime.Text = OrigApptTime
            End If

            Dim s = Split(Me.txtDate.Text, " ")
            Dim dt = s(0) & " 12:00:00 AM"
            Dim s2 = Split(Me.txtTime.Text, " ")
            Dim tm

            Try
                tm = "1/1/1900 " & s2(1) & s2(2)

            Catch ex As Exception
                tm = "1/1/1900 " & s2(0) & s2(1)
            End Try
            If frm.Name = "WCaller" Then
                Description = "Appt. Set from Warm Calling by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
            ElseIf frm.Name = "ConfirmingSingleRecord" Then
                Description = "Appt. Set by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
            ElseIf frm.Name = "Sales" Then
                If Me.Confirmed = True Then
                    Description = "Appt. Set from Sales Department by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & ". Confirmed by " & STATIC_VARIABLES.CurrentUser & ". (Confirming Bypassed)"
                Else
                    Description = "Appt. Set from Sales Department by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
                End If
            ElseIf frm.Name = "Recovery" Then
                Description = "Appt. Set from Recovery by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
            ElseIf frm.Name = "PreviousCustomer" Then
                Description = "Appt. Set from Previous Customer by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
            ElseIf frm.Name = "Finance" Then
                Description = "Appt. Set from Financing by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
            ElseIf frm.Name = "Installation" Then
                Description = "Appt. Set from Installation Department by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
            ElseIf frm.Name = "MarketingManager" Then
                If Me.Confirmed = True Then
                    Description = "Appt. Set from Marketing Manager by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & ". Confirmed by " & STATIC_VARIABLES.CurrentUser & ". (Confirming Bypassed)"
                Else
                    Description = "Appt. Set from Marketing Manager by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
                End If
            ElseIf frm.Name = "Administration" Then
                If Me.Confirmed = True Then
                    Description = "Appt. Set from Administration by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & ". Confirmed by " & STATIC_VARIABLES.CurrentUser & ". (Confirming Bypassed)"
                Else
                    Description = "Appt. Set from Administration by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
                End If
            ElseIf frm.Name = "Confirming" Then
                Description = "Appt. Set from Confirming by " & STATIC_VARIABLES.CurrentUser & ", set up with " & Me.cboSpokeWith.Text & " for " & day & ", " & CDate(dt).ToShortDateString & " at " & CDate(tm).ToShortTimeString & "."
            End If


            Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)

            Dim cmdIns As SqlCommand = New SqlCommand("dbo.LogSetAppt", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            Dim param2 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
            Dim param3 As SqlParameter = New SqlParameter("@Spokewith", Me.cboSpokeWith.Text)
            Dim param4 As SqlParameter = New SqlParameter("@ApptDate", dt)
            Dim param5 As SqlParameter = New SqlParameter("@Description", Description)
            Dim param6 As SqlParameter = New SqlParameter("@ApptTime", tm)
            Dim param7 As SqlParameter = New SqlParameter("@Notes", Me.RichTextBox1.Text)
            Dim param8 As SqlParameter = New SqlParameter("@Manager", STATIC_VARIABLES.CurrentManager)
            Dim param9 As SqlParameter = New SqlParameter("@Confirmed", Me.Confirmed)
            Dim param10 As SqlParameter = New SqlParameter("@Form", frm.Name.ToString)

            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cmdIns.Parameters.Add(param2)
            cmdIns.Parameters.Add(param3)
            cmdIns.Parameters.Add(param4)
            cmdIns.Parameters.Add(param5)
            cmdIns.Parameters.Add(param6)
            cmdIns.Parameters.Add(param7)
            cmdIns.Parameters.Add(param8)
            cmdIns.Parameters.Add(param9)
            cmdIns.Parameters.Add(param10)
            cnn.Open()
            Dim R1 As SqlDataReader
            R1 = cmdIns.ExecuteReader(CommandBehavior.CloseConnection)
            R1.Read()
            R1.Close()
            cnn.close()
            Me.Close()
            Me.Dispose()
        End If
        Select Case frm.Name
            Case "WCaller"
                If WCaller.Tab = "WC" Then
                    Dim i As Integer = WCaller.lvWarmCalling.Items.IndexOfKey(WCaller.lvWarmCalling.SelectedItems(0).Text)
                    WCaller.lvWarmCalling.SelectedItems(0).Remove()
                    WCaller.txtRecordsMatching.Text = CStr(CInt(WCaller.txtRecordsMatching.Text) - 1)
                    If WCaller.lvWarmCalling.Items.Count <> 0 Then
                        If i > WCaller.lvWarmCalling.Items.Count - 1 Then
                            WCaller.lvWarmCalling.Items(i - 1).Selected = True
                        Else
                            WCaller.lvWarmCalling.Items(i).Selected = True
                        End If
                    Else
                        Dim wc As New WarmCalling
                        wc.PullCustomerINFO("")
                    End If
                    Dim c As New WarmCalling.MyApptsTab.Populate(WCaller.cboFilter.Text)
                ElseIf WCaller.Tab = "MA" Then
                    Dim c As New WarmCalling.MyApptsTab.Populate(WCaller.cboFilter.Text)
                    Dim y As New WarmCalling
                    y.Populate()
                    WCaller.SuspendLayout()
                End If
            Case "Confirming"

                Dim c As New ConfirmingData
                c.PullCustomerINFO(Confirming.Tab, ID)
            Case "Installation"
            Case "Sales"
            Case "ConfirmerSingleRecord"
                Dim c As New CustomerHistory
                c.SetUp(frm, ID, ConfirmingSingleRecord.TScboCustomerHistory)
            Case "Administration"
                ''comeback
            Case "Finance"
                'comeback
            Case "Recovery"
                'comeback
            Case "PreviousCustomer"
                'comeback
            Case "ColdCalling"
                'comeback
            Case "SecondSource"
                'comeback
        End Select




    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub lblautonotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblautonotes.Click
        Me.cboautonotes.DroppedDown = True
    End Sub





    Private Sub lblSpokeWith_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSpokeWith.Click
        Me.cboSpokeWith.DroppedDown = True
    End Sub





    Private Sub txtDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.GotFocus
        Me.dtpApptDate.Focus()
    End Sub

    Private Sub dtpApptDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptDate.GotFocus
        Me.txtDate.Visible = False
        Me.txtDate.Text = Me.dtpApptDate.Value.ToString
    End Sub

    Private Sub dtpApptDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptDate.ValueChanged


        Me.txtDate.Text = Me.dtpApptDate.Value.ToString

    End Sub

    Private Sub dtpApptTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptTime.GotFocus
        Me.txtTime.Visible = False
        Me.txtTime.Text = Me.dtpApptTime.Value.ToString
    End Sub

    Private Sub dtpApptTime_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptTime.ValueChanged
        Me.txtTime.Text = Me.dtpApptTime.Value.ToString
    End Sub

    Private Sub txtTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime.GotFocus
        Me.dtpApptTime.Focus()


    End Sub
    Private Sub RichTextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.GotFocus
        Me.Label1.Visible = False
    End Sub

    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextBox1.LostFocus
        If Me.RichTextBox1.Text = "" Then
            Me.Label1.Visible = True
        End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Me.RichTextBox1.Focus()
    End Sub

    Private Sub cboSpokeWith_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSpokeWith.SelectedValueChanged
        If Me.cboSpokeWith.Text <> "" Then
            Me.lblSpokeWith.Text = Me.cboSpokeWith.Text
            Me.lblSpokeWith.ForeColor = Color.Black
            Me.lblSpokeWith.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        End If

    End Sub

End Class
