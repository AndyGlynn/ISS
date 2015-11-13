Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class CandCNotes
    Public Contact1 As String
    Public Contact2 As String
    Public ID As String
    Public OrigApptDate As String
    Public OrigApptTime As String
    Public frm As Form
    Dim d As Boolean = False

  

    Private Sub cboautonotes_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonotes.SelectedValueChanged
      Dim i As String
        Select Case Me.cboautonotes.Text
            Case "<Add New>"
                Me.cboautonotes.Text = Nothing

                i = InputBox$("Enter a new ""Auto Note"" here.", "Save Auto Note")

                If i = "" Then
                    MsgBox("You must enter Text to save this Auto Note!", MsgBoxStyle.Exclamation, "No Text Supplied")
                    Exit Sub
                End If
                Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)
                Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsCandCNotes", cnn)
                Dim cmdget As SqlCommand = New SqlCommand("dbo.GetCandCNotes", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@Notes", i)
                cmdINS.Parameters.Add(param1)
                cmdINS.CommandType = CommandType.StoredProcedure
                cmdget.CommandType = CommandType.StoredProcedure
                Dim r1 As SqlDataReader
                Dim r2 As SqlDataReader
                Me.cboautonotes.Items.Clear()
                Me.cboautonotes.Items.Add("<Add New>")
                Me.cboautonotes.Items.Add("___________________________________________________")
                cnn.Open()
                r2 = cmdINS.ExecuteReader
                r2.Close()
                r1 = cmdget.ExecuteReader
                While r1.Read
                    Me.cboautonotes.Items.Add(r1.Item(0))
                End While
                r1.Close()
                cnn.Close()

                Me.cboautonotes.Text = i
             
                Exit Select


            Case "___________________________________________________", ""
                Me.lblautonotes.ForeColor = Color.Gray
                Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Me.lblautonotes.Text = "Select Auto Note Here"
                Me.cboautonotes.SelectedItem = Nothing
                Exit Select
            Case Else
                Me.lblautonotes.Text = Me.cboautonotes.Text
                Me.lblautonotes.ForeColor = Color.Black
                Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                If Me.RichTextBox1.Text <> "" Then
                    Me.RichTextBox1.Text = Me.RichTextBox1.Text & " ," & Me.cboautonotes.Text
                Else
                    Me.RichTextBox1.Text = Me.cboautonotes.Text

                End If
               
        End Select

    End Sub



    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        If Me.dtpApptDate.Value < Today And Me.dtpApptDate.Enabled = True Then
            MsgBox("Appointment Date Must be for Today or Future!", MsgBoxStyle.Exclamation, "Cannot Move Appt.")
            Exit Sub
        End If
        Dim cnt As Integer = 0
        Me.ErrorProvider1.Clear()
        Me.ErrorProvider1.BlinkStyle = ErrorBlinkStyle.NeverBlink
        If Me.cboSpokeWith.Text = "" Then
            cnt += 1
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "You Must supply a Name!")
        End If

        If Me.RichTextBox1.Text = "" Then
            cnt += 1
            Me.ErrorProvider1.SetError(Me.RichTextBox1, "You must supply a reason for cancellation!")
        End If
        If Me.txtDate.Text = "" And Me.dtpApptDate.Enabled = True Then
            cnt += 1
            Me.ErrorProvider1.SetError(Me.dtpApptDate, "You must supply a new Appt. Date!")
        End If
        If Me.txtTime.Text = "" And Me.dtpApptTime.Enabled = True Then
            cnt += 1
            Me.ErrorProvider1.SetError(Me.dtpApptTime, "You must supply a new Appt. Time!")
        End If
        If cnt >= 1 Then
            Exit Sub
        End If

        Dim Description As String

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
        If Me.cboSpokeWith.Text = "Cancelled by Message" Then
            If Me.cbNoReschedule.Checked = True Then
                If Me.Contact2 <> "" Then
                    Description = Contact1 & " or " & Contact2 & " left a voicemail cancelling appt. with no interest in rescheduling."
                Else
                    Description = Contact1 & " left a voicemail cancelling appt. with no interest in rescheduling."
                End If
            ElseIf Me.cbCallBack.Checked = True Then
                If Me.Contact2 <> "" Then
                    Description = Contact1 & " or " & Contact2 & " left a voicemail cancelling appt., & would like a call back to reschedule."
                Else
                    Description = Contact1 & " left a voicemail cancelling appt., & would like a call back to reschedule."
                End If

            ElseIf Me.cbCallBack.Checked = False And Me.cbNoReschedule.Checked = False Then
                If Me.Contact2 <> "" Then
                    Description = Contact1 & " or " & Contact2 & " left a voicemail cancelling appt., and rescheduled for " & day & ", " & Me.dtpApptDate.Value.ToShortDateString & " at " & Me.dtpApptTime.Value.ToShortTimeString & "."
                Else
                    Description = Contact1 & " left a voicemail cancelling appt., and reschedule for " & day & ", " & Me.dtpApptDate.Value.ToShortDateString & " at " & Me.dtpApptTime.Value.ToShortTimeString & "."
                End If
            End If
        ElseIf Me.cboSpokeWith.Text <> "Cancelled by Message" Then
            If Me.cbNoReschedule.Checked = True Then
                Description = Me.cboSpokeWith.Text & " called in at " & Date.Now.ToShortTimeString & " spoke with " & STATIC_VARIABLES.CurrentUser & " to cancel Appt., and has no interest in rescheduling at this time."
            ElseIf Me.cbCallBack.Checked = True Then
                Description = Me.cboSpokeWith.Text & " called in at " & Date.Now.ToShortTimeString & " spoke with " & STATIC_VARIABLES.CurrentUser & " to cancel Appt., and would like a call back to reschedule."
            ElseIf Me.cbNoReschedule.Checked = False And Me.cbCallBack.Checked = False Then
                Description = Me.cboSpokeWith.Text & " called in at " & Date.Now.ToShortTimeString & " spoke with " & STATIC_VARIABLES.CurrentUser & " to cancel Appt., and rescheduled for " & day & ", " & Me.dtpApptDate.Value.ToShortDateString & " at " & Me.dtpApptTime.Value.ToShortTimeString & "."
            End If
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


        Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)

        Dim cmdIns As SqlCommand = New SqlCommand("dbo.LogCancelledAppt", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        Dim param2 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param3 As SqlParameter = New SqlParameter("@Spokewith", Me.cboSpokeWith.Text)
        Dim param4 As SqlParameter = New SqlParameter("@ApptDate", dt)
        Dim param5 As SqlParameter = New SqlParameter("@Description", Description)
        Dim param6 As SqlParameter = New SqlParameter("@ApptTime", tm)
        Dim param7 As SqlParameter = New SqlParameter("@Notes", Me.RichTextBox1.Text)
        cmdIns.CommandType = CommandType.StoredProcedure
        cmdIns.Parameters.Add(param1)
        cmdIns.Parameters.Add(param2)
        cmdIns.Parameters.Add(param3)
        cmdIns.Parameters.Add(param4)
        cmdIns.Parameters.Add(param5)
        cmdIns.Parameters.Add(param6)
        cmdIns.Parameters.Add(param7)
        cnn.Open()
        Dim R1 As SqlDataReader
        R1 = cmdIns.ExecuteReader(CommandBehavior.CloseConnection)
        R1.Read()
        R1.Close()
        cnn.close()




        Me.Close()
        Me.Dispose()
        Dim c As New ConfirmingData
        Dim c2 As New CustomerHistory
        Dim c3 As New WarmCalling
        If frm.Name = "Confirming" Then
            If Confirming.Tab = "Confirm" Then
                If Me.cbCallBack.Checked = True Or Me.cbNoReschedule.Checked = True Then
                    c.Populate(Confirming.Tab, Confirming.cboConfirmingPLS.Text, Confirming.cboConfirmingSLS.Text, Confirming.dpConfirming.Value.ToString, "Refresh")
                    c2.SetUp(frm, Confirming.lvConfirming.SelectedItems(0).Text, Confirming.TScboCustomerHistory)
                Else
                    c.Populate(Confirming.Tab, Confirming.cboConfirmingPLS.Text, Confirming.cboConfirmingSLS.Text, Confirming.dpConfirming.Value.ToString, "Populate")
                End If
            ElseIf Confirming.Tab = "Dispatch" Then
                If Me.cbCallBack.Checked = True Or Me.cbNoReschedule.Checked = True Then
                    c.Populate(Confirming.Tab, Confirming.cboSalesPLS.Text, Confirming.cboSalesSLS.Text, Confirming.dpSales.Value.ToString, "Refresh")
                    If Confirming.lvSales.SelectedItems.Count >= 1 Then
                        c2.SetUp(frm, Confirming.lvSales.SelectedItems(0).Text, Confirming.TScboCustomerHistory)
                    End If
                Else
                    c.Populate(Confirming.Tab, Confirming.cboSalesPLS.Text, Confirming.cboSalesSLS.Text, Confirming.dpSales.Value.ToString, "Populate")
                End If
            End If
        ElseIf frm.Name = "WCaller" Then
            If WCaller.Tab = "WC" Then
                c2.SetUp(frm, ID, WCaller.TScboCustomerHistory)
            Else
                If WCaller.lvMyAppts.SelectedItems(0).Group.Name <> "grpMemorized" Then
                    Dim x As New WarmCalling.MyApptsTab.Populate(WCaller.cboFilter.Text)
                    c3.Populate()
                Else
                    c2.SetUp(frm, ID, WCaller.TScboCustomerHistory)

                End If

            End If
        ElseIf frm.Name = "ConfirmingSingleRecord" Then
            c2.SetUp(frm, ID, ConfirmingSingleRecord.TScboCustomerHistory)

        ElseIf frm.Name = "Sales" Then
            If Sales.tbMain.SelectedIndex = 2 Then
                Dim x As New Issue_Leads(True, "")
            End If
            If Sales.tbMain.SelectedIndex = 1 Then
                c2.SetUp(frm, ID, Sales.TScboCustomerHistory)
            End If
        End If

        ''revisit for other forms 


    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub lblautonotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblautonotes.Click
        Me.cboautonotes.Focus()
        Me.cboautonotes.DroppedDown = True
    End Sub

    Private Sub cbCallBack_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCallBack.CheckedChanged
        If Me.cbCallBack.Checked = True Then
            Me.cbNoReschedule.Checked = False
            Me.dtpApptDate.Enabled = False
            Me.dtpApptTime.Enabled = False
            Me.txtDate.Enabled = False
            Me.txtTime.Enabled = False
            Me.txtDate.Text = ""
            Me.txtDate.Visible = True
            Me.txtTime.Text = ""
            Me.txtTime.Visible = True
        Else
            Me.dtpApptDate.Enabled = True
            Me.dtpApptTime.Enabled = True
            Me.txtDate.Enabled = True
            Me.txtTime.Enabled = True
            Me.txtDate.Text = ""
            Me.txtDate.Visible = True
            Me.txtTime.Text = ""
            Me.txtTime.Visible = True
        End If
    End Sub

    Private Sub cbNoReschedule_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbNoReschedule.CheckedChanged
        If Me.cbNoReschedule.Checked = True Then
            Me.cbCallBack.Checked = False
            Me.dtpApptDate.Enabled = False
            Me.dtpApptTime.Enabled = False
            Me.txtDate.Enabled = False
            Me.txtTime.Enabled = False
            Me.txtDate.Text = ""
            Me.txtDate.Visible = True
            Me.txtTime.Text = ""
            Me.txtTime.Visible = True
        Else
            Me.dtpApptDate.Enabled = True
            Me.dtpApptTime.Enabled = True
            Me.txtDate.Enabled = True
            Me.txtTime.Enabled = True
            Me.txtDate.Text = ""
            Me.txtDate.Visible = True
            Me.txtTime.Text = ""
            Me.txtTime.Visible = True
        End If
    End Sub

    Private Sub lblSpokeWith_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSpokeWith.Click
        'Me.cboSpokeWith.Focus()
        Me.cboSpokeWith.DroppedDown = True
    End Sub

   



  



 



    Private Sub txtDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.dtpApptDate.Select()
    End Sub

    Private Sub dtpApptDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.txtDate.Visible = False
        Me.txtDate.Text = Me.dtpApptDate.Value.ToString
    End Sub

    Private Sub dtpApptDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)


        Me.txtDate.Text = Me.dtpApptDate.Value.ToString

    End Sub

    Private Sub dtpApptTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.txtTime.Visible = False
        Me.txtTime.Text = Me.dtpApptTime.Value.ToString
    End Sub

    Private Sub dtpApptTime_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.txtTime.Text = Me.dtpApptTime.Value.ToString
    End Sub

    Private Sub txtTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.dtpApptTime.Select()


    End Sub

    Private Sub CandCNotes_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()

    End Sub

    Private Sub CandCNotes_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.cboSpokeWith.Items.Clear()

        Me.cboSpokeWith.Items.Add(Contact1)
        If Contact2 <> "" Then
            Me.cboSpokeWith.Items.Add(Contact2)
            Me.cboSpokeWith.Items.Add(Contact1 & " and " & Contact2)
        Else
            Me.cboSpokeWith.Text = Me.Contact1
        End If
        Me.cboSpokeWith.Items.Add("Cancelled by Message")
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
        Me.cbCallBack.Checked = False
        Me.cbNoReschedule.Checked = False
        Me.RichTextBox1.Text = ""

        Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdget As SqlCommand = New SqlCommand("dbo.GetCandCNotes", cnn)
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
        'Me.btnsave.Focus()

    End Sub

   
    

    Private Sub cboSpokeWith_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSpokeWith.SelectedValueChanged
        If Me.cboSpokeWith.Text <> "" Then
            Me.lblSpokeWith.Text = Me.cboSpokeWith.Text
            Me.lblSpokeWith.ForeColor = Color.Black
            Me.lblSpokeWith.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        End If
    End Sub

   

 
    Private Sub txtDate_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDate.GotFocus
        Me.dtpApptDate.Select()

    End Sub

    Private Sub txtTime_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTime.GotFocus
        Me.dtpApptTime.Select()
    End Sub

    Private Sub dtpApptDate_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptDate.GotFocus
        Me.txtDate.Visible = False
    End Sub

    Private Sub dtpApptTime_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpApptTime.GotFocus
        Me.txtTime.Visible = False
    End Sub


End Class
