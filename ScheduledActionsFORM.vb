Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient



Public Class ScheduledActions

    Dim panel As Panel
    Private Department As String
    Dim ttStatus As New ToolTip
    Dim ttNotes As New ToolTip
    Private notes As String
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)

    Public Structure SchedAction
        Public ID As String
        Public LeadNum As String
        Public Department As String
        Public AssignedTo As String
        Public ExecutionDate As String
        Public Notes As String
        Public AttachedFiles As Boolean
        Public SchedAction As String
        Public Completed As Boolean
        Public AttachedFilesHashValue As String
        Public Show As Boolean
        Public WhoCreated As String
    End Structure

    Public Sub SetUp(ByVal form As Form)


        Dim frm = form.Name



        If frm = "Sales" Then
            Department = "Sales"
            panel = Sales.pnlScheduledTasks
        ElseIf frm = "MarketingManager" Then
            Department = "Marketing"
            panel = MarketingManager.pnlScheduledTasks
        ElseIf frm = "Installation" Then
            Department = "Installation"
            panel = Installation.pnlScheduledTasks
        ElseIf frm = "Finance" Then
            Department = "Financing"
            panel = Finance.pnlScheduledTasks
        ElseIf frm = "Administration" Then
            Department = "Administration"
            panel = Administration.pnlScheduledTasks
        End If
        panel.Controls.Clear()



        Dim cmdGet As SqlCommand = New SqlCommand("dbo.PopScheduledActions", cnn)
        Dim cmdCNT As SqlCommand = New SqlCommand("dbo.CNTScheduledActions", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@Department", Department)
        Dim param2 As SqlParameter = New SqlParameter("@Department", Department)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdCNT.CommandType = CommandType.StoredProcedure

        cmdCNT.Parameters.Add(param1)
        Try


            cnn.Open()

            Dim sw As Boolean

            Dim r2 As SqlDataReader
            r2 = cmdCNT.ExecuteReader
            r2.Read()



            Dim o As Integer = CInt(r2.Item(0))
            If o Mod 2 = 0 Then
                sw = False
            Else
                sw = True
            End If

            r2.Close()








            cmdGet.Parameters.Add(param2)
            Dim R1 As SqlDataReader
            R1 = cmdGet.ExecuteReader





            While R1.Read


                Dim p As New Panel


                If sw = True Then
                    sw = False
                Else : sw = True
                End If

                Select Case sw

                    Case True
                        p.BackColor = Color.White
                        Exit Select
                    Case False
                        p.BackColor = Color.WhiteSmoke
                        Exit Select
                End Select




                p.Dock = DockStyle.Top
                p.BorderStyle = BorderStyle.None
                p.Size = New System.Drawing.Size(630, 28)
                'p.BackColor = Color.White
                p.Name = "pnl" & R1.Item(0)
                AddHandler p.MouseEnter, AddressOf scroll
                AddHandler p.MouseDown, AddressOf PanelControl


                panel.Controls.Add(p)


                Dim pctStatus As New PictureBox
                pctStatus.Size = New System.Drawing.Size(28, 28)
                pctStatus.Location = New System.Drawing.Point(4, 2)
                pctStatus.Name = "pctStatus" & R1.Item(0)
                pctStatus.Tag = R1.Item(1)
                pctStatus.Visible = False
                AddHandler pctStatus.Click, AddressOf Status
                AddHandler pctStatus.MouseDown, AddressOf Labels

                Dim lblDate As New Label
                lblDate.Text = R1.Item(2)
                lblDate.Name = "lblDate" & R1.Item(0)
                lblDate.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblDate.Location = New System.Drawing.Point(35, 7)
                lblDate.Size = New System.Drawing.Size(94, 16)
                lblDate.AutoEllipsis = True
                ttStatus.SetToolTip(lblDate, "This is the ""Due Date"" this task must be completed by")
                AddHandler lblDate.MouseDown, AddressOf Labels

                Dim lnk As New LinkLabel
                lnk.Text = R1.Item(3)
                lnk.Name = "lnk" & R1.Item(0)
                lnk.Location = New System.Drawing.Point(129, 7)
                lnk.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lnk.Size = New System.Drawing.Size(60, 16)
                AddHandler lnk.Click, AddressOf Link
                AddHandler lnk.MouseDown, AddressOf Labels


                Dim lblName As New Label

                Dim lblSA As New Label
                lblSA.Text = R1.Item(4)
                lblSA.Name = "lblSA" & R1.Item(0)
                lblSA.AutoEllipsis = True
                lblSA.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblSA.Location = New System.Drawing.Point(189, 7)
                ttStatus.SetToolTip(lblSA, "Description of Scheduled Action:" & vbCr & lblSA.Text)


                lblSA.Anchor = AnchorStyles.Left Or AnchorStyles.Right
                lblSA.AutoEllipsis = True

                AddHandler lblSA.MouseDown, AddressOf Labels



                lblName.Text = R1.Item(5)
                lblName.Name = "lblName" & R1.Item(0)
                'lblName.Dock = DockStyle.Right
                lblName.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Dim x = p.Size.Width
                lblName.Location = New System.Drawing.Point(x - 178, 7)
                lblName.Size = New System.Drawing.Size(120, 16)
                lblName.Anchor = AnchorStyles.Right
                lblName.AutoEllipsis = True
                ttStatus.SetToolTip(lblName, "This task has been assigned to: " & lblName.Text)
                AddHandler lblName.MouseDown, AddressOf Labels

                lblSA.Size = New System.Drawing.Size(lblName.Location.X - lblSA.Location.X - 5, 16)


                Dim pctNotes As New PictureBox
                pctNotes.Size = New System.Drawing.Size(28, 28)
                pctNotes.Location = New System.Drawing.Point(x - 58, 0)
                pctNotes.Name = "pctNotes" & R1.Item(0)
                pctNotes.Anchor = AnchorStyles.Right
                pctNotes.Tag = R1.Item(6)
                'pctNotes.Dock = DockStyle.Right
                ttNotes.SetToolTip(pctNotes, R1.Item(6))
                ttNotes.AutoPopDelay = 30000
                ttNotes.InitialDelay = 1
                ttNotes.ReshowDelay = 1
                pctNotes.Cursor = Cursors.Hand

                Dim c As New TruncateNotes
                c.Truncate(R1.Item(6), pctNotes)
                notes = c.NewSTRING

                ttNotes.SetToolTip(pctNotes, notes)
                pctNotes.Visible = False
                AddHandler pctNotes.Click, AddressOf show_notes
                AddHandler pctNotes.MouseDown, AddressOf Labels

                Dim pctAF As New PictureBox
                pctAF.Size = New System.Drawing.Size(28, 28)
                pctAF.Location = New System.Drawing.Point(x - 25, 0)
                pctAF.Name = "pctAF" & R1.Item(0)
                pctAF.Anchor = AnchorStyles.Right
                pctAF.Cursor = Cursors.Hand
                'pctAF.Dock = DockStyle.Right
                pctAF.Visible = False
                pctAF.Tag = R1.Item(7)
                Dim xx
                Dim fl = ""
                If pctAF.Tag <> "0" Then
                    Dim s = Split(pctAF.Tag, "\")
                    Dim index As Integer = 0
                    For Each xx In s
                        index += 1
                    Next

                    fl = s(index - 1)
                End If
             
                ttStatus.SetToolTip(pctAF, "This Icon is a Shortcut to an attached" & vbCr & "file needed to complete this task." & vbCr & "(Click icon to open file)" & vbCr & fl)
                AddHandler pctAF.Click, AddressOf OpenFile
                AddHandler pctAF.MouseDown, AddressOf Labels

                ''Add Controls




                p.Controls.Add(pctStatus)
                p.Controls.Add(lblDate)
                p.Controls.Add(lnk)
                p.Controls.Add(lblSA)

                p.Controls.Add(lblName)
                p.Controls.Add(pctNotes)
                p.Controls.Add(pctAF)
                '    ''Icon Logic
                If R1.Item(1) = True Then
                    pctStatus.Image = Main.ilScheduledTask.Images(0)
                    pctStatus.Visible = True
                    ttStatus.SetToolTip(pctStatus, "This Task has been Completed.")
                Else
                    If R1.Item(2) < Today Then
                        pctStatus.Image = Main.ilScheduledTask.Images(1)
                        pctStatus.Visible = True
                        ttStatus.SetToolTip(pctStatus, "This Task has not been completed" & vbCr & "& the due date has expired!")
                    End If
                End If
                If R1.Item(6) <> "" Then
                    pctNotes.Image = Main.ilScheduledTask.Images(2)
                    pctNotes.Visible = True
                End If
                If R1.Item(7) <> "0" Then
                    pctAF.Image = Main.ilScheduledTask.Images(3)
                    pctAF.Visible = True

                End If





            End While
            R1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
        End Try



    End Sub
    Dim i As ListViewItem
    Private Sub Link(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim frm As String = sender.findform.name
        Select Case frm
            Case "Sales"
                If Sales.cboSalesList.Text <> "Scheduled Tasks List" Then
                    Sales.cboSalesList.SelectedItem = "Scheduled Tasks List"
                End If
                For Each i In Sales.lvSales.Items
                    If i.Text = sender.text Then
                        i.Selected = True
                    End If
                Next
                If Sales.tbMain.SelectedIndex <> 1 Then
                    Sales.tbMain.SelectedIndex = 1
                End If
                If Sales.TabControl2.SelectedIndex <> 0 Then
                    Sales.TabControl2.SelectedIndex = 0
                End If
            Case "MarketingManager"
                ''comeback 

            Case "Installation"


            Case "Finance"


            Case "Administration"



        End Select



    End Sub
    Private Sub show_notes(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim pct As PictureBox = sender
        Me.ttNotes.Show(pct.Tag, pct, 30000)
    End Sub
 
    Private Sub scroll(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.panel.Select()
    End Sub
    Public Sub PanelControl(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'MsgBox(sender.ToString)
        Dim m = sender.findform
        Dim who As Panel = sender
        Dim selected As Boolean
        If who.BorderStyle = BorderStyle.FixedSingle And e.Button = MouseButtons.Left Then
            selected = True
        Else
            selected = False
        End If

        If who.BorderStyle = BorderStyle.FixedSingle Then
            who.BorderStyle = BorderStyle.None
        Else
            who.BorderStyle = BorderStyle.FixedSingle
        End If
        If m.name = "Sales" Then
            Dim c As Integer = Sales.pnlScheduledTasks.Controls.Count
            Dim i As Integer
            For i = 1 To c
                Dim all As Panel = Sales.pnlScheduledTasks.Controls(i - 1)
                If all.BorderStyle = BorderStyle.FixedSingle Then
                    all.BorderStyle = BorderStyle.None
                End If
            Next
            If selected = True Then
                who.BorderStyle = BorderStyle.None
            Else
                who.BorderStyle = BorderStyle.FixedSingle
            End If
            If who.Controls(0).Tag = "True" Then
                Sales.btnMarkTaskAsDone.Text = "Undo ""Mark Task as Done"""
                Sales.MarkTaskAsDoneToolStripMenuItem.Text = "Undo ""Mark Task as Done"""
                Sales.HideThisCompletedTaskToolStripMenuItem.Visible = True
                Sales.btnRemoveThisCompletedTask.Visible = True
            ElseIf who.Controls(0).Tag = "False" Then
                Sales.btnMarkTaskAsDone.Text = "Mark Task as Done"
                Sales.MarkTaskAsDoneToolStripMenuItem.Text = "Mark Task as Done"
                Sales.HideThisCompletedTaskToolStripMenuItem.Visible = False
                Sales.btnRemoveThisCompletedTask.Visible = False
            End If
            'Sales.btnScheduledTasks.ShowDropDown()

            'changed code
        ElseIf m.name = "MarketingManager" Then
            Dim c As Integer = MarketingManager.pnlScheduledTasks.Controls.Count
            Dim i As Integer
            For i = 1 To c
                Dim all As Panel = MarketingManager.pnlScheduledTasks.Controls(i - 1)
                If all.BorderStyle = BorderStyle.FixedSingle Then
                    all.BorderStyle = BorderStyle.None
                End If
            Next
            If selected = True Then
                who.BorderStyle = BorderStyle.None
            Else
                who.BorderStyle = BorderStyle.FixedSingle
            End If
            If who.Controls(0).Tag = "True" Then
                MarketingManager.btnMarkTaskAsDone.Text = "Undo ""Mark Task as Done"""
                MarketingManager.btnRemoveThisCompletedTask.Visible = True
            ElseIf who.Controls(0).Tag = "False" Then
                MarketingManager.btnMarkTaskAsDone.Text = "Mark Task as Done"
                MarketingManager.btnRemoveThisCompletedTask.Visible = False
            End If
            'MarketingManager.ScheduledTasks.ShowDropDown()
        ElseIf m.name = "Finance" Then
            Dim c As Integer = Finance.pnlScheduledTasks.Controls.Count
            Dim i As Integer
            For i = 1 To c
                Dim all As Panel = Finance.pnlScheduledTasks.Controls(i - 1)
                If all.BorderStyle = BorderStyle.FixedSingle Then
                    all.BorderStyle = BorderStyle.None
                End If
            Next
            If selected = True Then
                who.BorderStyle = BorderStyle.None
            Else
                who.BorderStyle = BorderStyle.FixedSingle
            End If
            If who.Controls(0).Tag = "True" Then
                Finance.btnMarkTaskAsDone.Text = "Undo ""Mark Task as Done"""
                Finance.btnRemoveThisCompletedTask.Visible = True
            ElseIf who.Controls(0).Tag = "False" Then
                Finance.btnEditScheduledTask.Text = "Mark Task as Done"
                Finance.btnRemoveThisCompletedTask.Visible = False
            End If
            'Finance.ScheduledTasks.ShowDropDown()
        ElseIf m.name = "Installation" Then
            Dim c As Integer = Installation.pnlScheduledTasks.Controls.Count
            Dim i As Integer
            For i = 1 To c
                Dim all As Panel = Installation.pnlScheduledTasks.Controls(i - 1)
                If all.BorderStyle = BorderStyle.FixedSingle Then
                    all.BorderStyle = BorderStyle.None
                End If
            Next
            If selected = True Then
                who.BorderStyle = BorderStyle.None
            Else
                who.BorderStyle = BorderStyle.FixedSingle
            End If
            If who.Controls(0).Tag = "True" Then
                Installation.btnMarkTaskAsDone.Text = "Undo ""Mark Task as Done"""
                Installation.btnRemoveThisCompletedTask.Visible = True
            ElseIf who.Controls(0).Tag = "False" Then
                Installation.btnMarkTaskAsDone.Text = "Mark Task as Done"
                Installation.btnRemoveThisCompletedTask.Visible = False
            End If
            'Installation.ScheduledTasks.ShowDropDown()
        ElseIf m.name = "Administation" Then
            Dim c As Integer = Administration.pnlScheduledTasks.Controls.Count
            Dim i As Integer
            For i = 1 To c
                Dim all As Panel = Administration.pnlScheduledTasks.Controls(i - 1)
                If all.BorderStyle = BorderStyle.FixedSingle Then
                    all.BorderStyle = BorderStyle.None
                End If
            Next
            If selected = True Then
                who.BorderStyle = BorderStyle.None
            Else
                who.BorderStyle = BorderStyle.FixedSingle
            End If
            If who.Controls(0).Tag = "True" Then
                Administration.btnMarkTaskAsDone.Text = "Undo ""Mark Task as Done"""
                Administration.btnRemoveThisCompletedTask.Visible = True
            ElseIf who.Controls(0).Tag = "False" Then
                Administration.btnMarkTaskAsDone.Text = "Mark Task as Done"
                Administration.btnRemoveThisCompletedTask.Visible = False
            End If
            'Administration.ScheduledTasks.ShowDropDown()
        End If
    End Sub
    Private Sub OpenFile(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim who As PictureBox = sender
        If InStr(who.Name, "pctAF") > 0 Then
            If who.Tag <> "" Then
                Try
                    Process.Start(who.Tag.ToString)
                Catch ex As Exception
                    MsgBox("The File Path is not valid or the file " & vbCr & "has been moved to another location!", MsgBoxStyle.Exclamation, "Cannot Locate File")
                End Try
            End If
        End If
    End Sub
    Private Sub Labels(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        Dim lbl As Control = sender

        PanelControl(lbl.Parent, e)
    End Sub
    Private Sub Status(ByVal Sender As System.Object, ByVal e As System.EventArgs)
        Dim lbl As PictureBox = Sender
        PanelControl(lbl.Parent, e)
    End Sub
    Public Sub Completed(ByVal ID As Integer)
        Try
            Dim cmdIns As SqlCommand = New SqlCommand("dbo.SAMarkAsDone", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
        End Try


    End Sub
    Public Sub UndoCompleted(ByVal ID As Integer)
        Try
            Dim cmdIns As SqlCommand = New SqlCommand("dbo.SAUndoMarkAsDone", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
        End Try
     

    End Sub
    Public Sub HideTask(ByVal ID As Integer)
        Try
            Dim cmdIns As SqlCommand = New SqlCommand("dbo.SAHide", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()

        Catch ex As Exception
            cnn.Close()
        End Try
   

    End Sub
    Public Sub HideAll(ByVal Form As Form)
        Try
            Dim cmdIns As SqlCommand = New SqlCommand("dbo.SAHideAll", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Department", Form.Name)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
        End Try

    End Sub
    Public Sub ShowAll(ByVal Form As Form)
        Try
            Dim cmdIns As SqlCommand = New SqlCommand("dbo.SAShowAll", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Department", Form.Name)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
        End Try

    End Sub
    Public Sub Preferences(ByVal y As Integer, ByVal Form As Form)
        Try
            Dim cmdIns As SqlCommand = New SqlCommand("dbo.SaUpdatePreferences", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Department", Form.Name)
            Dim param2 As SqlParameter = New SqlParameter("@num", y)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param2)
            cmdIns.Parameters.Add(param1)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
        End Try


    End Sub

    
End Class
