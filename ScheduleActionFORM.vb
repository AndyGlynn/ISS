
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System


Public Class ScheduleAction
    Public edit As Boolean = False
    Public EditId As String
    Public ep As New ErrorProvider
    Public RemoveErrP As Boolean = False
    Public tt As New ToolTip
    Public AtFile As Boolean = False
    Public Hash As String = "0"
    Public SAID As String

    Private Sub ScheduleAction_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    '' notes:
    '' font enumeration fontsytle.bolditalic = 3 in system enumeration
    '' hence the '3' instead of fontsytle.BoldAndItalic
    '' 
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.MdiParent = Main
        If edit = True Then
            Me.Populate_Edit()
            Exit Sub
        End If
        Me.txtLeadNumber.Text = STATIC_VARIABLES.CurrentID.ToString
        GetAutoCompleteSource()
        'GetScheduledAction()
        Dim a As New CUSTOMER_LABEL
        a.GetINFO(Me.txtLeadNumber.Text)
        Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, FontStyle.Bold, GraphicsUnit.Pixel, CType(0, Byte))
        Me.lblPhoneInfo.Text = a.Contact1Name & vbCrLf & a.StAddress & vbCrLf & vbCrLf & a.HousePhone & vbCrLf & a.AltPhone1 & "     " & a.AltPhone1Type & vbCrLf & a.AltPhone2 & "     " & a.AltPhone2Type
        'Me.lblContactInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))

        'Me.lblContactInfo.Text = "No Contact Information"
        If Me.txtLeadNumber.Text = "" Then
            Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"
        End If


    End Sub
    Private Sub Populate_Edit()

        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGETUSR As SqlCommand = New SqlCommand("SELECT * from iss.dbo.ScheduledTasks where id = '" & Me.EditId & "'", cnn)
        Dim ArUsers As New ArrayList
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdGETUSR.ExecuteReader
        While r1.Read
            Me.txtLeadNumber.Text = r1.Item(1)
            Dim a As New CUSTOMER_LABEL
            a.GetINFO(Me.txtLeadNumber.Text)
            Me.txtAssignedto.Text = r1.Item(3)
            Me.cboDept.SelectedItem = r1.Item(2)
            Dim x As SCHEDULE_ACTION_LOGIC
            'x.GetActionList(Me.cboDept.Text)
            Me.CboScheduledAction.SelectedItem = r1.Item(7)
            Me.dtpSA.Value = r1.Item(4)
            Me.rtfNotes.Text = r1.Item(5)
            If r1.Item(6) = True Then
                Me.Button3.Text = "Remove File..."
                Me.Hash = r1.Item(9)
                Me.AtFile = True
            End If
            Me.Button1.Text = "Edit"
            Me.Text = "Edit Scheduled Task"
        End While
        r1.Close()
        cnn.Close()

    End Sub
    Private Sub txtLeadNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLeadNumber.KeyPress
        'tt = New ToolTip
        tt.RemoveAll()
        Select Case e.KeyChar
            Case "/"c, "\"c, "!"c, "#"c, "$"c, "%"c, "^"c, "&"c, "("c, ")"c, "?"c, "<"c, ">"c, "@"c
                tt.Show("!@#$%^&*()?<>\/  Are illegal characters to instert.", Me.txtLeadNumber, 3000)
                Me.txtLeadNumber.Text = ""
                Exit Select
        End Select
    End Sub


    Private Sub txtLeadNumber_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadNumber.TextChanged
        Dim str As String = ""
        str = Me.txtLeadNumber.Text
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
                Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, FontStyle.Bold, GraphicsUnit.Pixel, CType(0, Byte))
                Me.lblPhoneInfo.Text = a.Contact1Name & vbCrLf & a.StAddress & vbCrLf & vbCrLf & a.HousePhone & vbCrLf & a.AltPhone1 & "     " & a.AltPhone1Type & vbCrLf & a.AltPhone2 & "     " & a.AltPhone2Type
                Me.ep.Clear()
            End If
        End If
    End Sub
    Private Sub GetAutoCompleteSource()
        Me.txtAssignedto.AutoCompleteCustomSource.Clear()
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
                Me.txtAssignedto.AutoCompleteCustomSource.Add("Admin")
            ElseIf ArUsers(g).ToString <> "Admin" Then
                Me.txtAssignedto.AutoCompleteCustomSource.Add(ArUsers(g).ToString)
            End If
        Next
    End Sub
    Private Sub GetScheduledAction()
        Me.CboScheduledAction.Items.Clear()
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT ScheduledAction from iss.dbo.actionlist", cnn)
        Dim r1 As SqlDataReader
        cnn.Open()
        r1 = cmdGET.ExecuteReader
        Dim ArActions As New ArrayList
        While r1.Read
            ArActions.Add(r1.Item(0))
        End While
        r1.Close()
        cnn.Close()
        Dim g As Integer = 0
        Me.CboScheduledAction.Items.Add("<Add New>")
        Me.CboScheduledAction.Items.Add("____________________________")
        Me.CboScheduledAction.Items.Add("")

        For g = 0 To ArActions.Count - 1
            Me.CboScheduledAction.Items.Add(ArActions(g).ToString)
        Next

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
                    ep.SetError(Me.txtLeadNumber, "Invalid Lead Number")
                    Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, FontStyle.Bold, GraphicsUnit.Pixel, CType(0, Byte))
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Me.Button3.Text.Contains("Remove") Then
            Me.remove_file()
          
        Else
            Me.AttachFile1()
        End If


    End Sub
    Private Sub remove_file()
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGETUSR As SqlCommand = New SqlCommand("update iss.dbo.ScheduledTasks set attachedfiles = 'False', attachedhashvalue = '0' where id = '" & Me.EditId & "'", cnn)
        Try
            cnn.Open()
            cmdGETUSR.ExecuteNonQuery()
            cnn.Close()
            System.IO.Directory.Delete(STATIC_VARIABLES.SAAttachedFileDirectory & Me.EditId, True)
            Me.Hash = "0"
            Me.AtFile = False
            Me.Button3.Text = "Attach a File..."
        Catch ex As Exception
            cnn.Close()
        End Try
     
    End Sub

    Private Sub CboScheduledAction_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboScheduledAction.DropDown
        If Me.CboScheduledAction.Items.Count = 0 Then
            Me.tt.Show("You Must Select a Department to" & vbCr & "Populate Schelduled Actions!", Me.CboScheduledAction, 5000)
        Else
            Me.tt.RemoveAll()
        End If
    End Sub

  

    Private Sub CboScheduledAction_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboScheduledAction.SelectedValueChanged
        Dim x As String = ""
        x = Me.CboScheduledAction.Text
        Select Case x
            Case Is = "<Add New>"
                Dim str As String = ""
                str = InputBox$("Enter new Action to add for the " & Me.cboDept.Text & " department." & vbCr & "(Actions should be a short description of a standard action that is recurring & recyclable. If you need to supply more detail please use the notes field.", "New Action")
                If str.ToString.Length < 2 Then
                    Exit Sub
                End If
                Dim g As New SCHEDULE_ACTION_LOGIC
                g.InsertNewAction(Me.cboDept.Text, str)
                'Me.GetScheduledAction()
                If str <> "" Then
                    Me.CboScheduledAction.SelectedItem = str.ToString
                Else
                    Me.CboScheduledAction.SelectedItem = ""
                End If

                Exit Select
            Case Is = ""
                Exit Select
            Case Is = "________________________"
                Me.CboScheduledAction.Text = ""
                Exit Select
                'Case Else
                '    Dim sal As New SCHEDULE_ACTION_LOGIC
                '    sal.GetActionList(Me.saCboDepart.Text)
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim leadNum As String = ""
        Dim ExecDate As String = ""
        Dim Notes As String = ""
        Dim AssignedTo As String = ""
        Dim Department As String = ""
        Dim SchedAction As String = ""
        leadNum = Me.txtLeadNumber.Text
        If leadNum = "" Then
            Exit Sub
        End If
        ExecDate = Me.dtpSA.Value.Date
        Notes = Me.rtfNotes.Text
        AssignedTo = Me.txtAssignedto.Text
        Department = Me.cboDept.Text
        SchedAction = Me.CboScheduledAction.Text
        Me.ErrorProvider1.Clear()
        Dim cnt As Integer = 0
        If ExecDate = "" Then
            Me.ErrorProvider1.SetError(Me.dtpSA, "Required Field")
            cnt += 1
            'Exit Sub
        End If
        If AssignedTo = "" Then
            Me.ErrorProvider1.SetError(Me.txtAssignedto, "Required Field")
            cnt += 1
            'Exit Sub
        End If
        If Department = "" Then
            Me.ErrorProvider1.SetError(Me.cboDept, "Required Field")
            cnt += 1
            'Exit Sub
        End If
        If SchedAction = "" Then
            Me.ErrorProvider1.SetError(Me.CboScheduledAction, "Required Field")
            cnt += 1
        End If
        If cnt >= 1 Then
            Exit Sub
        End If
        '' do not use cmraude
        '' use stored procedure method instead
      
        Dim x As New SCHEDULE_ACTION_LOGIC.InsertSA
        Dim hash As String = ""
        Dim aFile As String = ""
        If AtFile = True Then
            hash = Me.Hash
            aFile = "1"
        End If
        If AtFile = False Then
            hash = "0"
            aFile = "0"
        End If
        If Me.Button1.Text.Contains("Edit") Then
            x.UpdateSchedAction(Me.EditId, leadNum, Department, AssignedTo, ExecDate, Notes, AtFile, SchedAction, hash)
        Else
            x.InsertNewSchedAction(leadNum, Department, AssignedTo, ExecDate, Notes, AtFile, SchedAction, hash, False)
        End If

        If aFile = "1" Then
            Me.AttachFile2()
        End If

        Me.txtLeadNumber.Text = ""
        Me.txtAssignedto.Text = ""
        Me.rtfNotes.Text = ""
        Me.CboScheduledAction.Items.Clear()
        Me.CboScheduledAction.Text = ""
        Me.cboDept.Items.Clear()
        Me.cboDept.Text = ""
        Me.dtpSA.Value = Date.Today.Date
        Me.AtFile = False
        Me.Close()
    End Sub
    Dim path As String = STATIC_VARIABLES.SAAttachedFileDirectory
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public Sub AttachFile1()

        Dim opfd As New Windows.Forms.OpenFileDialog
        opfd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop

        opfd.ShowDialog()
        Dim d As String

        d = opfd.FileName.ToString

        If d.ToString = "" Then

            Me.txtLeadNumber.Select()
            Exit Sub
        End If
        Me.Hash = d
        Me.AtFile = True


    End Sub
    Public Sub AttachFile2()
        Dim d = Me.Hash
        Dim e
        Dim cnt As Integer = 0
        Dim ch As Char
        If d <> "0" Then
            For Each ch In d
                If ch = "\" Then
                    cnt += 1
                End If
            Next
        Else
            Exit Sub
        End If
        Dim id = Me.SAID


        e = Split(d, "\", cnt + 1)

        Dim z
        z = e(cnt)
        Dim sfp As String = Replace(d.ToString, "\" & z, "")
        '''' create directory 
        Dim file As String = path
        System.IO.Directory.CreateDirectory(path + ID.ToString)
        file = file + ID.ToString + "\" + z.ToString
        Hash = file
      

    

        'Dim x
        'Dim icnt As Integer = 0

     

        Try
            System.IO.File.Copy(d, path + ID.ToString & "\" & z.ToString)
        Catch ex As Exception
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")
        End Try
        Dim x As New SCHEDULE_ACTION_LOGIC.InsertSA
        x.Update_AF_path(Hash, SAID)
      
    
    End Sub

    Private Sub cboDept_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDept.SelectedValueChanged
        Dim x As New SCHEDULE_ACTION_LOGIC
        x.GetActionList(Me.cboDept.Text)
    End Sub
End Class
