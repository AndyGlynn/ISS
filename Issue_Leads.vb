Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Drawing.Graphics
Imports system.Globalization


Public Class Issue_Leads
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Private cnn1 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Private cnn2 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Dim y As Panel = Sales.pnlIssue
    Dim ttStatus As New ToolTip
    Dim ttNotes As New ToolTip
    Dim ttlabels As New ToolTip
    Dim ttNotes2 As New ToolTip
    Dim ttReps As New ToolTip
    Dim w As Integer = y.Width
    Dim cm As ContextMenuStrip = Sales.cmIssue
    Dim Rep2 As Boolean = False
    Public panelsize As String
    Public arReps As List(Of USER_LOGICv2.Employee)



 
    Public Sub New(ByVal setup As Boolean, ByVal growshrink As String)
        Dim leadCnt As Integer = 0
        If setup = False Then
            Me.panelsize = growshrink
            Me.ReLayout()
            Exit Sub
        End If

        Try
            y.Visible = False
            y.Controls.Clear()
            If Sales.tpIssueLeads.Controls.Count >= 2 Then
                Dim headerpanel As Panel = Sales.tpIssueLeads.Controls.Item(1)
                Sales.tpIssueLeads.Controls.Remove(headerpanel)
            End If
          

            Dim d As String = Sales.dtpIssueLeads.Value
            Dim s = Split(d, " ")
            d = s(0) & " 00:00:00.000"
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetIssued", cnn)
            Dim cmdCNT As SqlCommand = New SqlCommand("Select Count (ID) from Enterlead where ApptDate = '" & d & "' and MarketingResults = 'Confirmed'", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@dt", d)
            cmdGet.CommandType = CommandType.StoredProcedure
            cmdGet.Parameters.Add(param1)
            cnn.Open()


            Dim r2 As SqlDataReader = cmdCNT.ExecuteReader
            r2.Read()
            Dim header As New Panel
            header.Dock = DockStyle.Top
            header.BorderStyle = BorderStyle.None
            header.BackColor = System.Drawing.SystemColors.GradientActiveCaption
            header.Name = "MainHeader"
            header.Size = New Size(w, 48)
            Sales.tpIssueLeads.Controls.Add(header)
            Sales.pnlIssue.BringToFront()
            Dim lblheader As New Label
            lblheader.AutoSize = True
            lblheader.Font = New System.Drawing.Font("Verdana", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            lblheader.Location = New System.Drawing.Point(12, 12)
            lblheader.Name = "lblheader"
            header.Controls.Add(lblheader)

            If r2.Item(0) <= 0 Then
                lblheader.Text = "There are no Appointments to Issue for " & day(Sales.dtpIssueLeads.Value.DayOfWeek) & ", " & Sales.dtpIssueLeads.Value.ToShortDateString
                y.Visible = True
                Exit Sub
            Else
                leadCnt = CType(r2.Item(0), Integer)
                lblheader.Text = "Appointments for " & day(Sales.dtpIssueLeads.Value.DayOfWeek) & ", " & Sales.dtpIssueLeads.Value.ToShortDateString
                Dim lblApptCnt As New Label
                Dim lblCancelCnt As New Label
                lblApptCnt.Name = "lblApptCnt"
                lblApptCnt.AutoSize = True
                lblApptCnt.Font = New System.Drawing.Font("Verdana", 11.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblApptCnt.Location = New System.Drawing.Point(header.Width - 240, 3)
                lblApptCnt.Text = "Appointments to Issue- " & r2.Item(0)
                lblApptCnt.Anchor = AnchorStyles.Right
                header.Controls.Add(lblApptCnt)
                lblCancelCnt.Name = "lblCancelCnt"
                lblCancelCnt.AutoSize = True
                lblCancelCnt.Font = New System.Drawing.Font("Verdana", 11.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblCancelCnt.Location = New System.Drawing.Point(header.Width - 240, 27)
                lblCancelCnt.Anchor = AnchorStyles.Right
                header.Controls.Add(lblCancelCnt)
            End If
            Dim sw As Boolean
            Dim o As Integer = CInt(r2.Item(0))
            If o Mod 2 = 0 Then
                sw = True
            Else
                sw = False
            End If
            r2.Close()
            Dim R1 As SqlDataReader
            R1 = cmdGet.ExecuteReader
            Dim cnt As Integer = 0
            Dim cnt2 As Integer = 0
            While R1.Read
                Dim tm As DateTime = R1.Item(22)
                Dim time As String
                time = (tm.ToString("g", _
                  CultureInfo.CreateSpecificCulture("en-us")))

                Dim z = Split(time, " ")
                time = z(1) & " " & z(2)
                cnt += 1
                Dim first As Boolean
                Dim create As Boolean
                If create = False Then
                    cnt2 += 1
                End If
                Dim compare As String
                If cnt = 1 Then
                    compare = time
                    create = True  ''Creates First Time Header 
                    first = True
                End If

                If time <> compare Then
                    create = True
                    compare = time

                End If
                Dim p As New Panel
                If create = True Then
                    Dim a As New Panel  ''Create Header Panel 
                    a.Dock = DockStyle.Top
                    a.BorderStyle = BorderStyle.None
                    a.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
                    a.Name = "header" & time
                    a.Size = New Size(w, 25)
                    'a.Location = New System.Drawing.Point(0, 0)
                    a.Margin = New Padding(0)
                    a.Tag = time

                    y.Controls.Add(a)


                    Dim lblH As New Label  '' Create Header Label 
                    lblH.AutoSize = True
                    lblH.Font = New System.Drawing.Font("Verdana", 11.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblH.Location = New System.Drawing.Point(3, 3)
                    lblH.Name = "lblH" & time
                    lblH.Text = time & " Appointments - "


                    a.Controls.Add(lblH)
                    y.Controls.SetChildIndex(a, 0)
                    AddHandler a.MouseEnter, AddressOf scroll

                    create = False
                    'MsgBox("Header Added")
                End If
                cnt2 += 1





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
                p.Name = "pnl" & R1.Item(0)
                p.Size = New Size(w, 48)
                p.Margin = New Padding(0)
                p.Tag = time
                p.ContextMenuStrip = cm
                p.Cursor = Cursors.Hand

                AddHandler p.MouseEnter, AddressOf scroll
                AddHandler p.MouseDown, AddressOf PanelControl
                'p.Location = New System.Drawing.Point(0, 0)

                y.Controls.Add(p)
                y.Controls.SetChildIndex(p, 0)

                'p.Location = New System.Drawing.Point(0, t)
                't = t + p.Size.Height

                'MsgBox(p.Location.X)

                ttStatus.AutoPopDelay = 30000

                Dim pctIcon As New PictureBox
                pctIcon.Size = New System.Drawing.Size(32, 32)
                pctIcon.Location = New System.Drawing.Point(6, 8)
                pctIcon.Name = "pctIcon" & R1.Item(0)
                pctIcon.Cursor = Cursors.Hand
                pctIcon.Anchor = AnchorStyles.Left
                AddHandler pctIcon.Click, AddressOf show_notes


                If R1.Item(21) = "Not Hit" Or R1.Item(21) = "Not Issued" Then
                    pctIcon.Tag = "Exclaim"
                    ttStatus.SetToolTip(pctIcon, "This Appointment was " & R1.Item(21) & vbCr & " the last time it was Confirmed" & vbCr & "There is a risk we may lose this" & vbCr & "potential customer if we don't" & vbCr & "make it today!")
                End If
                If R1.Item(20) = True Then
                    pctIcon.Tag = "PC"
                    ttStatus.SetToolTip(pctIcon, "Previous Customer")
                End If
                If R1.Item(19) = True Then
                    pctIcon.Tag = "Recovery"
                    ttStatus.SetToolTip(pctIcon, "Recovery Appointment")
                End If
                ttStatus.ToolTipIcon = ToolTipIcon.Info
                ttStatus.ToolTipTitle = "Appointment Status:"




                If pctIcon.Tag = "" Then
                    Try
                        Dim cmdNEW As SqlCommand = New SqlCommand("dbo.IsNewIssued", cnn1)
                        Dim param2 As SqlParameter = New SqlParameter("@id", R1.Item(0))

                        cmdNEW.CommandType = CommandType.StoredProcedure
                        cmdNEW.Parameters.Add(param2)

                        cnn1.Open()
                        Dim r3 As SqlDataReader = cmdNEW.ExecuteReader
                        While r3.Read()
                            If r3.Item(0) >= 1 And R1.Item(21) = "" Then
                                pctIcon.Tag = "New"
                                ttStatus.SetToolTip(pctIcon, "New Appointment: This is the first time this" & vbCr & "Appt has been through the Sales Dept." & vbCr & _
                                "to be issued. Appt Generated on " & R1.Item(23))
                            End If
                        End While
                        r3.Close()
                        cnn1.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If


                If pctIcon.Tag = "Recovery" Then
                    pctIcon.Image = Sales.ImgIssued3.Images(0)
                ElseIf pctIcon.Tag = "PC" Then
                    pctIcon.Image = Sales.ImgIssued3.Images(1)
                ElseIf pctIcon.Tag = "Exclaim" Then
                    pctIcon.Image = Sales.ImgIssued3.Images(2)
                ElseIf pctIcon.Tag = "New" Then
                    pctIcon.Image = Sales.ImgIssued3.Images(3)
                Else
                    pctIcon.Visible = False
                End If

                Dim lnk As New LinkLabel
                lnk.Text = R1.Item(0)
                lnk.Name = "lnk" & R1.Item(0)
                lnk.Location = New System.Drawing.Point(44, 8)
                lnk.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lnk.Size = New System.Drawing.Size(78, 32)
                lnk.AutoSize = False

                lnk.TextAlign = ContentAlignment.MiddleLeft
                lnk.Anchor = AnchorStyles.Left
                AddHandler lnk.Click, AddressOf Link
                'AddHandler lnk.MouseDown, AddressOf Labels
                Dim name As String
                Try
                    If R1.Item(2) = R1.Item(4) Then
                        name = R1.Item(1) & " and " & R1.Item(3) & " " & R1.Item(2)
                    ElseIf R1.Item(4) = "" And R1.Item(3) <> "" Then
                        name = R1.Item(1) & " and " & R1.Item(3) & " " & R1.Item(2)
                    ElseIf R1.Item(3) = "" Then
                        name = R1.Item(1) & " " & R1.Item(2)
                    ElseIf R1.Item(2) <> R1.Item(4) Then
                        name = R1.Item(1) & " " & R1.Item(2) & " and " & R1.Item(3) & " " & R1.Item(4)
                    End If
                Catch ex As Exception
                    MsgBox(ex)
                End Try

                Dim fs1 As Single = 9.75
                Dim strsize As New System.Drawing.SizeF()
                strsize = Me.MeasureString(name, New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, _
                                                    System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                Dim lblName As New Label
                lblName.Text = name
                lblName.Name = "lblName" & R1.Item(0)

                lblName.Location = New System.Drawing.Point(128, 0)
                Dim x As Integer = (p.Size.Width - (390)) * 0.44
                lblName.Size = New System.Drawing.Size(x, p.Height)



                lblName.Anchor = AnchorStyles.Right
                lblName.AutoEllipsis = False
                lblName.AutoSize = False


                lblName.TextAlign = ContentAlignment.MiddleCenter
                lblName.Anchor = AnchorStyles.Left
                lblName.Cursor = Cursors.Hand

                ttlabels.SetToolTip(lblName, "Customer(s) Name")
                AddHandler lblName.MouseDown, AddressOf Labels
                Dim adText As String = R1.Item(5) & vbCr & R1.Item(6) & ", " & R1.Item(7) & " " & R1.Item(8)
                Dim lblAddress As New Label

                Dim addy As String = R1.Item(5) & " " & R1.Item(6) & ", " & R1.Item(7) & " " & R1.Item(8)
                strsize = Me.MeasureString(addy, New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, _
                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                lblAddress.Name = "lblAddress" & R1.Item(0)

                x = lblName.Width + 134
                lblAddress.Location = New System.Drawing.Point(x, 0)
                x = (p.Size.Width - (390)) * 0.3
                lblAddress.Size = New System.Drawing.Size(x, p.Height)



                If strsize.Width >= lblAddress.Size.Width Then
                    lblAddress.Text = adText
                    lblAddress.Tag = addy
                Else
                    lblAddress.Text = addy
                    lblAddress.Tag = adText
                End If
                lblAddress.Anchor = AnchorStyles.Right
                lblAddress.AutoEllipsis = True
                lblAddress.AutoSize = False
                lblAddress.TextAlign = ContentAlignment.MiddleCenter
                lblAddress.Cursor = Cursors.Hand

                lblAddress.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right Or AnchorStyles.Bottom Or AnchorStyles.Top), System.Windows.Forms.AnchorStyles)
                ttlabels.SetToolTip(lblAddress, "Appointment Address")
                AddHandler lblAddress.MouseDown, AddressOf Labels

                Dim Prod As String
                Dim prod2 As String
                If R1.Item(10) = "" And R1.Item(11) = "" Then
                    Prod = R1.Item(9)
                ElseIf R1.Item(10) <> "" And R1.Item(11) = "" Then
                    Prod = R1.Item(9) & " - " & R1.Item(10)
                ElseIf R1.Item(10) <> "" And R1.Item(11) <> "" Then
                    Prod = R1.Item(9) & " - " & R1.Item(10) & " - " & R1.Item(11)
                End If
                Dim lblProduct As New Label
                ttlabels.SetToolTip(lblProduct, "Product(s): " & Prod)
                strsize = Me.MeasureString(Prod, New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, _
                System.Drawing.GraphicsUnit.Point, CType(0, Byte)))





                Dim pctSI As New PictureBox

                lblProduct.Name = "lblProduct" & R1.Item(0)

                x = lblName.Width + 134 + lblAddress.Width + 6
                lblProduct.Location = New System.Drawing.Point(x, 0)

                lblProduct.Anchor = AnchorStyles.Right
                lblProduct.Cursor = Cursors.Hand
                Dim acro1 As String
                Dim acro2 As String
                Dim acro3 As String
                Try
                    acro1 = R1.Item(12)
                Catch ex As Exception
                    acro1 = ""
                End Try
                Try
                    acro2 = R1.Item(13)
                Catch ex As Exception
                    acro2 = ""
                End Try
                Try
                    acro3 = R1.Item(14)
                Catch ex As Exception
                    acro3 = ""
                End Try
                If acro2 = "" And acro3 = "" Then
                    prod2 = acro1
                ElseIf acro2 <> "" And acro3 = "" Then
                    prod2 = acro1 & " - " & acro2
                ElseIf acro2 <> "" And acro3 <> "" Then
                    prod2 = acro1 & " - " & acro2 & " - " & acro3
                End If
                If acro1 = "  " Then
                    prod2 = ""
                End If

                lblProduct.AutoEllipsis = True
                lblProduct.AutoSize = False
                lblProduct.TextAlign = ContentAlignment.MiddleCenter




                AddHandler lblProduct.MouseDown, AddressOf Labels


                Dim notes As String = R1.Item(15)
                If notes <> "" Then
                    Dim c As New TruncateNotes
                    c.Truncate(notes, pctSI)
                    notes = c.NewSTRING
                    ttNotes.SetToolTip(pctSI, notes)
                    AddHandler pctSI.Click, AddressOf show_notes

                    pctSI.Image = Sales.ImgIssued3.Images(4)
                    pctSI.Visible = True
                Else
                    pctSI.Visible = False

                End If

                pctSI.Size = New System.Drawing.Size(32, 32)
                x = p.Width - 286
                pctSI.Location = New System.Drawing.Point(x, 8)
                pctSI.Name = "pctSI" & R1.Item(0)
                x = (p.Size.Width - ((lblAddress.Location.X + lblAddress.Width) + (p.Width - pctSI.Location.X + 16)))
                lblProduct.Size = New System.Drawing.Size(x, p.Height)






                If (strsize.Width) >= lblProduct.Width Then
                    lblProduct.Text = prod2
                    lblProduct.Tag = Prod
                Else
                    lblProduct.Text = Prod
                    lblProduct.Tag = prod2
                End If




                lblProduct.Font = New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblName.Font = New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblAddress.Font = New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))



                ttNotes.ToolTipIcon = ToolTipIcon.Info
                ttNotes.ToolTipTitle = "Special Instructions"
                ttNotes.AutoPopDelay = 30000
                ttNotes.InitialDelay = 1
                ttNotes.ReshowDelay = 1
                pctSI.Cursor = Cursors.Hand
                pctSI.Anchor = AnchorStyles.Right
                pctSI.BringToFront()

                Dim cboRep1 As New ComboBox
                cboRep1.Name = "cboRep1" & R1.Item(0)
                cboRep1.DropDownStyle = ComboBoxStyle.DropDownList
                cboRep1.Size = New Size(163, 21)
                x = p.Width - 241
                cboRep1.Location = New System.Drawing.Point(x, 14)
                cboRep1.Anchor = AnchorStyles.Right
                cboRep1.Tag = R1.Item(0)
                cboRep1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Pop_Reps(cboRep1)
                cboRep1.SelectedItem = R1.Item(16)
                If R1.Item(16) <> "" Then
                    If cboRep1.SelectedItem = Nothing Or cboRep1.SelectedItem = "" Then
                        cboRep1.Items.Add(R1.Item(16))
                        cboRep1.SelectedItem = R1.Item(16)
                    End If
                End If
                AddHandler cboRep1.SelectedValueChanged, AddressOf Cbo
                ''look for rep in list if not there add it & make it selected value 

                Dim cboRep2 As New ComboBox
                cboRep2.Name = "cboRep2" & R1.Item(0)
                cboRep2.DropDownStyle = ComboBoxStyle.DropDownList
                cboRep2.Size = New Size(163, 21)
                cboRep2.Location = New System.Drawing.Point(x, 14)
                cboRep2.Visible = False
                cboRep2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                cboRep2.Anchor = AnchorStyles.Right
                cboRep2.Tag = R1.Item(0)
                Pop_Reps(cboRep2)
                If R1.Item(17) <> "" Then
                    cboRep2.SelectedItem = R1.Item(17)
                    If cboRep2.SelectedItem = Nothing Or cboRep2.SelectedItem = "" Then
                        cboRep2.Items.Add(R1.Item(17))
                        cboRep2.SelectedItem = R1.Item(17)
                    End If
                    Rep2 = True
                End If

                AddHandler cboRep2.SelectedValueChanged, AddressOf Cbo
                ''look for rep in list if not there add it & make it selected value 

                Dim pctRep2 As New PictureBox
                pctRep2.Size = New System.Drawing.Size(16, 16)
                x = p.Width - 72
                pctRep2.Location = New System.Drawing.Point(x, 16)
                pctRep2.Name = "pctRep2" & R1.Item(0)
                pctRep2.Tag = "Add"
                pctRep2.Image = Sales.ImgIssued.Images(5)
                pctRep2.Cursor = Cursors.Hand
                pctRep2.Anchor = AnchorStyles.Right
                ttlabels.SetToolTip(pctRep2, "Add a 2nd Rep")
                AddHandler pctRep2.Click, AddressOf Add_Rep

                Dim Mnotes As String
                Dim pctMN As New PictureBox
                pctMN.Anchor = AnchorStyles.Right
                Try

                    Mnotes = R1.Item(18)
                Catch ex As Exception
                    Mnotes = ""
                End Try

                If Mnotes <> "" Then
                    Dim c As New TruncateNotes
                    c.Truncate(Mnotes, pctMN)
                    Mnotes = c.NewSTRING
                    ttNotes2.SetToolTip(pctMN, Mnotes)

                    AddHandler pctMN.Click, AddressOf show_notes

                    pctMN.Image = Sales.ImgIssued3.Images(6)
                    pctMN.Visible = True
                Else
                    pctMN.Visible = False

                End If

                pctMN.Size = New System.Drawing.Size(32, 32)
                x = p.Width - 46
                pctMN.Location = New System.Drawing.Point(x, 8)
                pctMN.Name = "pctMN" & R1.Item(0)



                ttNotes2.ToolTipIcon = ToolTipIcon.Info
                ttNotes2.ToolTipTitle = "Manager Notes"
                ttNotes2.AutoPopDelay = 30000
                ttNotes2.InitialDelay = 1
                ttNotes2.ReshowDelay = 1
                pctMN.Cursor = Cursors.Hand
                pctMN.BringToFront()


                p.Controls.Add(pctIcon)
                p.Controls.Add(lblName)
                p.Controls.Add(lnk)
                p.Controls.Add(lblAddress)
                p.Controls.Add(lblProduct)
                p.Controls.Add(pctSI)
                p.Controls.Add(cboRep1)
                p.Controls.Add(cboRep2)
                p.Controls.Add(pctRep2)
                p.Controls.Add(pctMN)

                If Rep2 = True Then
                    Dim i As PictureBox = pctRep2
                    p.Size = New Size(p.Width, 62)
                    p.Controls.Item(6).Location = New System.Drawing.Point(p.Controls.Item(6).Location.X, 7)
                    p.Controls.Item(7).Location = New System.Drawing.Point(p.Controls.Item(7).Location.X, 34)
                    p.Controls.Item(7).Visible = True
                    i.Location = New System.Drawing.Point(i.Location.X, 9)
                    i.Image = Sales.ImgIssued.Images(7)
                    i.Tag = "minus"
                    ttlabels.SetToolTip(i, "Close 2nd Rep")
                    Dim h
                    For h = 0 To p.Controls.Count - 1
                        Dim Cntl As Control = p.Controls.Item(h)
                        If Cntl.Name.Contains("lbl") Then
                            Cntl.Size = New Size(Cntl.Width, p.Height)
                            Cntl.Location = New System.Drawing.Point(Cntl.Location.X, 0)
                        End If
                    Next
                    Rep2 = False
                End If

            End While
            R1.Close()
            cnn.Close()

            Dim time2 As String
            Dim counter As Integer = 0
            For e As Integer = 0 To y.Controls.Count - 1
                If e = 0 Then
                    time2 = y.Controls.Item(0).Tag
                End If
                If counter = 0 And e <> 0 Then
                    time2 = y.Controls.Item(e).Tag
                End If
                If y.Controls.Item(e).BackColor <> System.Drawing.SystemColors.GradientInactiveCaption And y.Controls.Item(e).Tag = time2 Then
                    time2 = y.Controls.Item(e).Tag
                    counter += 1
                Else
                    y.Controls.Item(e).Controls.Item(0).Text = y.Controls.Item(e).Controls.Item(0).Text & counter
                    counter = 0
                End If
            Next ' Gets Count of # of Appts. Matching Header and adds count to header 

            '' Add called & Cancelled Appt. to the bottom of the issue leads panel 
            ''Create Header Panel 
            Dim cmdCancels As SqlCommand = New SqlCommand("dbo.GetIssuedCancels", cnn)

            Dim paramC As SqlParameter = New SqlParameter("@dt", d)
            cmdCancels.CommandType = CommandType.StoredProcedure
            cmdCancels.Parameters.Add(paramC)
            cnn.Open()
            Dim r As SqlDataReader = cmdCancels.ExecuteReader

            Dim ccCnt As Integer = 0
            While r.Read
                ccCnt += 1
                If ccCnt = 1 Then

                    Dim cc As New Panel '' Create Cancelled Appts Header Label 
                    cc.Dock = DockStyle.Top
                    cc.BorderStyle = BorderStyle.None
                    cc.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
                    cc.Name = "CCHeader"
                    cc.Size = New Size(w, 25)




                    y.Controls.Add(cc)


                    Dim lblcc As New Label
                    lblcc.AutoSize = True
                    lblcc.Font = New System.Drawing.Font("Verdana", 11.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblcc.Location = New System.Drawing.Point(3, 3)
                    lblcc.Name = "lblcc"
                    lblcc.Text = "Called and Cancelled Appointments - "


                    cc.Controls.Add(lblcc)
                    y.Controls.SetChildIndex(cc, 0)
                    AddHandler cc.MouseEnter, AddressOf scroll
                End If
                If sw = True Then
                    sw = False
                Else : sw = True
                End If
                Dim pc As New Panel
                Select Case sw

                    Case True
                        pc.BackColor = Color.White
                        Exit Select
                    Case False
                        pc.BackColor = Color.WhiteSmoke
                        Exit Select
                End Select

                Dim tm As DateTime = r.Item(22)
                Dim time As String
                time = (tm.ToString("g", _
                  CultureInfo.CreateSpecificCulture("en-us")))

                Dim z = Split(time, " ")
                time = z(1) & " " & z(2)

                pc.Dock = DockStyle.Top
                pc.BorderStyle = BorderStyle.None
                pc.Name = "pnl" & r.Item(0)
                pc.Size = New Size(w, 48)
                pc.Margin = New Padding(0)
                pc.Tag = "CC"
                pc.ContextMenuStrip = cm
                pc.Cursor = Cursors.Hand

                AddHandler pc.MouseEnter, AddressOf scroll
                AddHandler pc.MouseDown, AddressOf PanelControl


                y.Controls.Add(pc)
                y.Controls.SetChildIndex(pc, 0)

                Dim pctIcon As New PictureBox
                pctIcon.Size = New System.Drawing.Size(32, 32)
                pctIcon.Location = New System.Drawing.Point(6, 8)
                pctIcon.Name = "pctIcon" & r.Item(0)
                pctIcon.Cursor = Cursors.Hand
                pctIcon.Anchor = AnchorStyles.Left
                pctIcon.Image = Sales.ImgIssued3.Images(8)
                AddHandler pctIcon.Click, AddressOf show_notes
                ttStatus.SetToolTip(pctIcon, "Customer Called & Cancelled" & vbCr & "Appt. Time: " & time)

                Dim lnk As New LinkLabel
                lnk.Text = r.Item(0)
                lnk.Name = "lnk" & r.Item(0)
                lnk.Location = New System.Drawing.Point(44, 8)
                lnk.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lnk.Size = New System.Drawing.Size(78, 32)
                lnk.AutoSize = False
                lnk.TextAlign = ContentAlignment.MiddleLeft
                lnk.Anchor = AnchorStyles.Left
                AddHandler lnk.Click, AddressOf Link

                Dim name As String
                Try
                    If r.Item(2) = r.Item(4) Then
                        name = r.Item(1) & " and " & r.Item(3) & " " & r.Item(2)
                    ElseIf r.Item(4) = "" And r.Item(3) <> "" Then
                        name = r.Item(1) & " and " & r.Item(3) & " " & r.Item(2)
                    ElseIf r.Item(3) = "" Then
                        name = r.Item(1) & " " & r.Item(2)
                    ElseIf r.Item(2) <> r.Item(4) Then
                        name = r.Item(1) & " " & r.Item(2) & " and " & r.Item(3) & " " & r.Item(4)
                    End If
                Catch ex As Exception
                    MsgBox(ex)
                End Try

                Dim fs1 As Single = 9.75
                Dim strsize As New System.Drawing.SizeF()
                strsize = Me.MeasureString(name, New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, _
                                                    System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                Dim lblName As New Label
                lblName.Text = name
                lblName.Name = "lblName" & r.Item(0)
                lblName.Location = New System.Drawing.Point(128, 0)
                Dim x As Integer = (pc.Size.Width - (390)) * 0.44
                lblName.Size = New System.Drawing.Size(x, pc.Height)
                lblName.Anchor = AnchorStyles.Right
                lblName.AutoEllipsis = False
                lblName.AutoSize = False
                lblName.TextAlign = ContentAlignment.MiddleCenter
                lblName.Anchor = AnchorStyles.Left
                lblName.Cursor = Cursors.Hand
                ttlabels.SetToolTip(lblName, "Customer(s) Name")
                AddHandler lblName.MouseDown, AddressOf Labels

                Dim lblAddress As New Label
                Dim adText As String = r.Item(5) & vbCr & r.Item(6) & ", " & r.Item(7) & " " & r.Item(8)
                Dim addy As String = r.Item(5) & " " & r.Item(6) & ", " & r.Item(7) & " " & r.Item(8)
                strsize = Me.MeasureString(addy, New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, _
                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                lblAddress.Name = "lblAddress" & r.Item(0)

                x = lblName.Width + 134
                lblAddress.Location = New System.Drawing.Point(x, 0)
                x = (pc.Size.Width - (390)) * 0.3
                lblAddress.Size = New System.Drawing.Size(x, pc.Height)



                If strsize.Width >= lblAddress.Size.Width Then
                    lblAddress.Text = adText
                    lblAddress.Tag = addy
                Else
                    lblAddress.Text = addy
                    lblAddress.Tag = adText
                End If
                lblAddress.Anchor = AnchorStyles.Right
                lblAddress.AutoEllipsis = True
                lblAddress.AutoSize = False
                lblAddress.TextAlign = ContentAlignment.MiddleCenter
                lblAddress.Cursor = Cursors.Hand

                lblAddress.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right Or AnchorStyles.Bottom Or AnchorStyles.Top), System.Windows.Forms.AnchorStyles)
                ttlabels.SetToolTip(lblAddress, "Appointment Address")
                AddHandler lblAddress.MouseDown, AddressOf Labels

                Dim Prod As String
                Dim prod2 As String
                If r.Item(10) = "" And r.Item(11) = "" Then
                    Prod = r.Item(9)
                ElseIf r.Item(10) <> "" And r.Item(11) = "" Then
                    Prod = r.Item(9) & " - " & r.Item(10)
                ElseIf r.Item(10) <> "" And r.Item(11) <> "" Then
                    Prod = r.Item(9) & " - " & r.Item(10) & " - " & r.Item(11)
                End If
                Dim lblProduct As New Label
                ttlabels.SetToolTip(lblProduct, "Product(s): " & Prod)
                strsize = Me.MeasureString(Prod, New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, _
                System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                lblProduct.Name = "lblProduct" & r.Item(0)
                x = lblName.Width + 134 + lblAddress.Width + 6
                lblProduct.Location = New System.Drawing.Point(x, 0)
                lblProduct.Anchor = AnchorStyles.Right
                lblProduct.Cursor = Cursors.Hand
                Dim acro1 As String
                Dim acro2 As String
                Dim acro3 As String
                Try
                    acro1 = r.Item(12)
                Catch ex As Exception
                    acro1 = ""
                End Try
                Try
                    acro2 = r.Item(13)
                Catch ex As Exception
                    acro2 = ""
                End Try
                Try
                    acro3 = r.Item(14)
                Catch ex As Exception
                    acro3 = ""
                End Try
                If acro2 = "" And acro3 = "" Then
                    prod2 = acro1
                ElseIf acro2 <> "" And acro3 = "" Then
                    prod2 = acro1 & " - " & acro2
                ElseIf acro2 <> "" And acro3 <> "" Then
                    prod2 = acro1 & " - " & acro2 & " - " & acro3
                End If
                lblProduct.AutoEllipsis = True
                lblProduct.AutoSize = False
                lblProduct.TextAlign = ContentAlignment.MiddleCenter
                AddHandler lblProduct.MouseDown, AddressOf Labels

                Dim pctSI As New PictureBox
                Dim notes As String = r.Item(15)
                If notes <> "" Then
                    Dim c As New TruncateNotes
                    c.Truncate(notes, pctSI)
                    notes = c.NewSTRING
                    ttNotes.SetToolTip(pctSI, notes)
                    AddHandler pctSI.Click, AddressOf show_notes

                    pctSI.Image = Sales.ImgIssued3.Images(4)
                    pctSI.Visible = True
                Else
                    pctSI.Visible = False
                End If

                pctSI.Size = New System.Drawing.Size(32, 32)
                x = pc.Width - 286
                pctSI.Location = New System.Drawing.Point(x, 8)
                pctSI.Name = "pctSI" & r.Item(0)
                x = (pc.Size.Width - ((lblAddress.Location.X + lblAddress.Width) + (pc.Width - pctSI.Location.X + 16)))
                lblProduct.Size = New System.Drawing.Size(x, pc.Height)






                If (strsize.Width) >= lblProduct.Width Then
                    lblProduct.Text = prod2
                    lblProduct.Tag = Prod
                Else
                    lblProduct.Text = Prod
                    lblProduct.Tag = prod2
                End If




                lblProduct.Font = New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblName.Font = New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblAddress.Font = New System.Drawing.Font("Verdana", fs1!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))



                ttNotes.ToolTipIcon = ToolTipIcon.Info
                ttNotes.ToolTipTitle = "Special Instructions"
                ttNotes.AutoPopDelay = 30000
                ttNotes.InitialDelay = 1
                ttNotes.ReshowDelay = 1
                pctSI.Cursor = Cursors.Hand
                pctSI.Anchor = AnchorStyles.Right
                pctSI.BringToFront()

                Dim cboRep1 As New ComboBox
                cboRep1.Name = "cboRep1" & r.Item(0)
                cboRep1.DropDownStyle = ComboBoxStyle.DropDownList
                cboRep1.Size = New Size(163, 21)
                x = pc.Width - 241
                cboRep1.Location = New System.Drawing.Point(x, 14)
                cboRep1.Anchor = AnchorStyles.Right
                cboRep1.Tag = r.Item(0)
                cboRep1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Pop_Reps(cboRep1)
                cboRep1.SelectedItem = r.Item(16)
                If r.Item(16) <> "" Then
                    If cboRep1.SelectedItem = Nothing Or cboRep1.SelectedItem = "" Then
                        cboRep1.Items.Add(r.Item(16))
                        cboRep1.SelectedItem = r.Item(16)
                    End If
                End If
                AddHandler cboRep1.SelectedValueChanged, AddressOf Cbo
                ''look for rep in list if not there add it & make it selected value 

                Dim cboRep2 As New ComboBox
                cboRep2.Name = "cboRep2" & r.Item(0)
                cboRep2.DropDownStyle = ComboBoxStyle.DropDownList
                cboRep2.Size = New Size(163, 21)
                cboRep2.Location = New System.Drawing.Point(x, 14)
                cboRep2.Visible = False
                cboRep2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                cboRep2.Anchor = AnchorStyles.Right
                cboRep2.Tag = r.Item(0)
                Pop_Reps(cboRep2)
                If r.Item(17) <> "" Then
                    cboRep2.SelectedItem = r.Item(17)
                    If cboRep2.SelectedItem = Nothing Or cboRep2.SelectedItem = "" Then
                        cboRep2.Items.Add(r.Item(17))
                        cboRep2.SelectedItem = r.Item(17)
                    End If
                    Rep2 = True
                End If

                AddHandler cboRep2.SelectedValueChanged, AddressOf Cbo
                ''look for rep in list if not there add it & make it selected value 

                Dim pctRep2 As New PictureBox
                pctRep2.Size = New System.Drawing.Size(16, 16)
                x = pc.Width - 72
                pctRep2.Location = New System.Drawing.Point(x, 16)
                pctRep2.Name = "pctRep2" & r.Item(0)
                pctRep2.Tag = "Add"
                pctRep2.Image = Sales.ImgIssued.Images(5)
                pctRep2.Cursor = Cursors.Hand
                pctRep2.Anchor = AnchorStyles.Right
                ttlabels.SetToolTip(pctRep2, "Add a 2nd Rep")
                AddHandler pctRep2.Click, AddressOf Add_Rep

                Dim Mnotes As String
                Dim pctMN As New PictureBox
                pctMN.Anchor = AnchorStyles.Right
                Try
                    Mnotes = r.Item(18)
                Catch ex As Exception
                    Mnotes = ""
                End Try

                If Mnotes <> "" Then
                    Dim c As New TruncateNotes
                    c.Truncate(Mnotes, pctMN)
                    Mnotes = c.NewSTRING
                    ttNotes2.SetToolTip(pctMN, Mnotes)

                    AddHandler pctMN.Click, AddressOf show_notes

                    pctMN.Image = Sales.ImgIssued3.Images(6)
                    pctMN.Visible = True
                Else
                    pctMN.Visible = False
                End If

                pctMN.Size = New System.Drawing.Size(32, 32)
                x = pc.Width - 46
                pctMN.Location = New System.Drawing.Point(x, 8)
                pctMN.Name = "pctMN" & r.Item(0)



                ttNotes2.ToolTipIcon = ToolTipIcon.Info
                ttNotes2.ToolTipTitle = "Manager Notes"
                ttNotes2.AutoPopDelay = 30000
                ttNotes2.InitialDelay = 1
                ttNotes2.ReshowDelay = 1
                pctMN.Cursor = Cursors.Hand
                pctMN.BringToFront()


                pc.Controls.Add(pctIcon)
                pc.Controls.Add(lblName)
                pc.Controls.Add(lnk)
                pc.Controls.Add(lblAddress)
                pc.Controls.Add(lblProduct)
                pc.Controls.Add(pctSI)
                pc.Controls.Add(cboRep1)
                pc.Controls.Add(cboRep2)
                pc.Controls.Add(pctRep2)
                pc.Controls.Add(pctMN)
                If Rep2 = True Then
                    Dim i As PictureBox = pctRep2
                    pc.Size = New Size(pc.Width, 62)
                    pc.Controls.Item(6).Location = New System.Drawing.Point(pc.Controls.Item(6).Location.X, 7)
                    pc.Controls.Item(7).Location = New System.Drawing.Point(pc.Controls.Item(7).Location.X, 34)
                    pc.Controls.Item(7).Visible = True
                    i.Location = New System.Drawing.Point(i.Location.X, 9)
                    i.Image = Sales.ImgIssued.Images(7)
                    i.Tag = "minus"
                    ttlabels.SetToolTip(i, "Close 2nd Rep")
                    Dim h
                    For h = 0 To pc.Controls.Count - 1
                        Dim Cntl As Control = pc.Controls.Item(h)
                        If Cntl.Name.Contains("lbl") Then
                            Cntl.Size = New Size(Cntl.Width, pc.Height)
                            Cntl.Location = New System.Drawing.Point(Cntl.Location.X, 0)
                        End If
                    Next
                    Rep2 = False
                End If
           
            End While
            r.Close()
            cnn.Close()

            Dim cancelled As Integer = 0
            For e As Integer = 0 To y.Controls.Count - 1
                If y.Controls.Item(e).Tag = "CC" Then
                    cancelled += 1
                End If
                If y.Controls.Item(e).Controls.Item(0).Text.Contains("Called and Cancelled Appointments - ") Then
                    y.Controls.Item(e).Controls.Item(0).Text = y.Controls.Item(e).Controls.Item(0).Text & cancelled.ToString
                    Sales.tpIssueLeads.Controls.Item(Sales.tpIssueLeads.Controls.Count - 1).Controls.Item(2).Text = "Cancelled Appointments- " & cancelled.ToString
                    Exit For
                Else
                    Sales.tpIssueLeads.Controls.Item(Sales.tpIssueLeads.Controls.Count - 1).Controls.Item(2).Text = "Cancelled Appointments- " & cancelled.ToString
                    ''Add realign here 
                End If
            Next ' Gets Count of Cancelled Appts & adds to Cancelled Header

            ''Adds appt counts and cancelled counts to main header 



            '' This part upsizes fonts & labels to fill extra space when loaded in maximized mode (Same code as the Relayout subprocedure 
            '' that shrinks or grows on size change of form 

            Dim fs As Single = 9.75
            Dim fcnt As Integer
            Dim doit As Integer
            Dim buffer As Integer = 10
            For e As Integer = 0 To y.Controls.Count - 1
                If y.Controls.Item(e).Name.Contains("pnl") Then
                    doit += 1

                    Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                    Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                    Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                    Dim strsize As New System.Drawing.SizeF()
                    strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                          System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                    If strsize.Width <= (lbladdress.Width - buffer) Then
                        fcnt += 1

                    End If
                    strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                       System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                    If strsize.Width <= (lblname.Width - buffer) Then
                        fcnt += 1
                    End If
                    strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                             System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                    If strsize.Width <= (lblproduct.Width - buffer) Then
                        fcnt += 1
                    End If

                End If
            Next

            If fcnt = (doit * 3) Then
                fcnt = 0
                doit = 0
                fs = fs + 1
                For e As Integer = 0 To y.Controls.Count - 1
                    If y.Controls.Item(e).Name.Contains("pnl") Then
                        doit += 1

                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                        Dim strsize As New System.Drawing.SizeF()
                        strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                        If strsize.Width <= (lbladdress.Width - buffer) Then
                            fcnt += 1

                        End If
                        strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblname.Width - buffer) Then
                            fcnt += 1
                        End If
                        strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                                 System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblproduct.Width - buffer) Then
                            fcnt += 1
                        End If

                    End If
                Next

            End If

            If fcnt = (doit * 3) Then

                fs = fs + 1
                For e As Integer = 0 To y.Controls.Count - 1
                    If y.Controls.Item(e).Name.Contains("pnl") Then
                        doit += 1

                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                        Dim strsize As New System.Drawing.SizeF()
                        strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                        If strsize.Width <= (lbladdress.Width - buffer) Then
                            fcnt += 1

                        End If
                        strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblname.Width - buffer) Then
                            fcnt += 1
                        End If
                        strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                                 System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblproduct.Width - buffer) Then
                            fcnt += 1
                        End If

                    End If
                Next
            End If



            If fs > 9.75 Then
                For e As Integer = 0 To y.Controls.Count - 1

                    If y.Controls.Item(e).Name.Contains("pnl") Then


                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                        Dim lnk As LinkLabel = y.Controls.Item(e).Controls.Item(2)
                        lblname.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lbladdress.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lblproduct.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        lnk.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    End If
                Next
            End If












            y.Visible = True
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub ReLayout()

        Dim bORs As Boolean
        Dim fs As Single
        Dim Ofs As Single
        If Sales.pnlIssue.Controls.Count = 0 Then
            Exit Sub
        Else
            Try
                fs = Sales.pnlIssue.Controls.Item(0).Controls.Item(1).Font.Size
            Catch ex As Exception
                Dim err As String = ex.Message
                Exit Sub
            End Try
        End If
        Dim fcnt As Integer
        Dim doit As Integer
        Dim buffer As Integer = 10
        Dim q As Integer



        Try

            For q = 0 To Sales.pnlIssue.Controls.Count - 1

                Dim y As Panel = Sales.pnlIssue.Controls.Item(q)
                If y.Name.Contains("pnl") Then
                    Dim lblname As Label = y.Controls.Item(1)
                    Dim lbladdress As Label = y.Controls.Item(3)
                    Dim lblproduct As Label = y.Controls.Item(4)
                    Dim strsize As New System.Drawing.SizeF()
                    Dim strsize2 As New System.Drawing.SizeF()
                    Dim pctSI = y.Controls.Item(5)
                    Dim x As Integer = (lblname.Parent.Size.Width - (390)) * 0.44

                    lblname.Size = New System.Drawing.Size(x, y.Height)
                    x = lblname.Width + 134
                    lbladdress.Location = New System.Drawing.Point(x, 0)
                    x = (y.Size.Width - (390)) * 0.3
                    lbladdress.Size = New System.Drawing.Size(x, y.Height)
                    strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                          System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                    strsize2 = Me.MeasureString(lbladdress.Tag, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                          System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                    Dim tag = lbladdress.Tag
                    Dim txt = lbladdress.Text
                    Dim longest As String
                    If strsize.Width > strsize2.Width Then
                        longest = "txt"
                    Else
                        longest = "tag"
                    End If
                    If longest = "txt" Then
                        If strsize.Width >= lbladdress.Width Then
                            lbladdress.Text = tag
                            lbladdress.Tag = txt
                        Else
                            lbladdress.Text = txt
                            lbladdress.Tag = tag
                        End If
                    Else  ''longest = tag 
                        If strsize2.Width >= lbladdress.Width Then
                            lbladdress.Text = txt
                            lbladdress.Tag = tag
                        Else
                            lbladdress.Text = tag
                            lbladdress.Tag = txt
                        End If
                    End If
                    x = lblname.Width + 134 + lbladdress.Width + 6
                    lblproduct.Location = New System.Drawing.Point(x, 0)
                    x = (y.Size.Width - ((lbladdress.Location.X + lbladdress.Width) + (y.Width - pctSI.Location.X + 16)))
                    lblproduct.Size = New System.Drawing.Size(x, y.Height)
                    strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                    strsize2 = Me.MeasureString(lblproduct.Tag, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                          System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                    tag = lblproduct.Tag
                    txt = lblproduct.Text
                    If strsize.Width > strsize2.Width Then
                        longest = "txt"
                    Else
                        longest = "tag"
                    End If
                    If longest = "txt" Then
                        If strsize.Width >= lblproduct.Width Then
                            lblproduct.Text = tag
                            lblproduct.Tag = txt
                        Else
                            lblproduct.Text = txt
                            lblproduct.Tag = tag
                        End If
                    Else  ''longest = tag 
                        If strsize2.Width >= lblproduct.Width Then
                            lblproduct.Text = txt
                            lblproduct.Tag = tag
                        Else
                            lblproduct.Text = tag
                            lblproduct.Tag = txt
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


        If Me.panelsize = "shrink" Then

            Try

                If y.Controls.Count <> 0 Then
                    fs = y.Controls.Item(0).Controls.Item(1).Font.Size
                End If
                For e As Integer = 0 To y.Controls.Count - 1
                    If y.Controls.Item(e).Name.Contains("pnl") Then
                        doit += 1

                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)

                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                        Dim strsize As New System.Drawing.SizeF()
                        strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                        If strsize.Width >= (lbladdress.Width - buffer) Then
                            fcnt += 1
                            'MsgBox(lbladdress.Name & " strsize.width = " & strsize.Width & " lbladdress.width - buffer = " & (lbladdress.Width - buffer))

                        End If
                        strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width >= (lblname.Width - buffer) Then
                            fcnt += 1


                        End If
                        'MsgBox(lblname.Name & " strsize.width = " & strsize.Width & " lblname.width - buffer = " & (lblname.Width - buffer))
                        strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                                 System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width >= (lblproduct.Width - buffer) Then
                            fcnt += 1
                            '[MsgBox(lblproduct.Name & " strsize.width = " & strsize.Width & " lblproduct.width - buffer = " & (lblproduct.Width - buffer))

                        End If

                    End If
                Next
            Catch ex As Exception
                'MsgBox(ex.ToString) ''took message box because of known error, when the controls active scrollbar it cause a size chenge before creation done 
                '' Doesnt matter though cuz after creation in full screen it does another Relayout anyway Cant think of a way to kick it out when this happens 
                '' so i just commented out MSGBOX
            End Try





            If fcnt > 0 Then  ''10.75 check +-

                If fs > 10.75 Then
                    fs = fs - 1
                End If

                fcnt = 0
                doit = 0





                For e As Integer = 0 To y.Controls.Count - 1
                    If y.Controls.Item(e).Name.Contains("pnl") Then
                        doit += 1

                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)

                        Dim strsize As New System.Drawing.SizeF()
                        strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                        If strsize.Width >= (lbladdress.Width - buffer) Then
                            fcnt += 1


                        End If
                        strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width >= (lblname.Width - buffer) Then
                            fcnt += 1

                        End If
                        strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                                 System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width >= (lblproduct.Width - buffer) Then
                            fcnt += 1

                        End If

                    End If
                Next

            End If

            If fcnt > 0 Then
                If fs > 9.75 Then
                    fs = fs - 1
                End If





                For e As Integer = 0 To y.Controls.Count - 1
                    If y.Controls.Item(e).Name.Contains("pnl") Then
                        doit += 1

                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                        Dim strsize As New System.Drawing.SizeF()
                        strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                        If strsize.Width >= (lbladdress.Width - buffer) Then
                            fcnt += 1
                        Else
                            fcnt -= 1
                        End If
                        strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width >= (lblname.Width - buffer) Then
                            fcnt += 1

                        End If
                        strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                                 System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width >= (lblproduct.Width - buffer) Then
                            fcnt += 1

                        End If

                    End If
                Next

            End If

            For e As Integer = 0 To y.Controls.Count - 1

                If y.Controls.Item(e).Name.Contains("pnl") Then


                    Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                    Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                    Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                    Dim lnk As LinkLabel = y.Controls.Item(e).Controls.Item(2)
                    lblname.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lbladdress.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblproduct.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lnk.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                End If
            Next

        End If
        fcnt = 0
        doit = 0



        y = Sales.pnlIssue
        If Me.panelsize = "grow" Then
            If y.Controls.Count <> 0 Then
                Try
                    fs = y.Controls.Item(0).Controls.Item(1).Font.Size
                Catch ex As Exception
                    Dim err As String = ex.Message
                End Try
            End If
            For e As Integer = 0 To y.Controls.Count - 1
                If y.Controls.Item(e).Name.Contains("pnl") Then
                    doit += 1

                    Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                    Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                    Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)

                    Dim strsize As New System.Drawing.SizeF()
                    strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                          System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                    If strsize.Width <= (lbladdress.Width - buffer) Then
                        fcnt += 1


                    End If
                    strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                       System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                    If strsize.Width <= (lblname.Width - buffer) Then
                        fcnt += 1


                    End If
                    strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                             System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                    If strsize.Width <= (lblproduct.Width - buffer) Then
                        fcnt += 1


                    End If

                End If
            Next







            If fcnt = (doit * 3) Then  ''10.75 check +-
                If fs < 10.75 Then
                    fs = fs + 1
                End If

                fcnt = 0
                doit = 0





                For e As Integer = 0 To y.Controls.Count - 1
                    If y.Controls.Item(e).Name.Contains("pnl") Then
                        doit += 1

                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                        Dim strsize As New System.Drawing.SizeF()
                        strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                        If strsize.Width <= (lbladdress.Width - buffer) Then
                            fcnt += 1


                        End If
                        strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblname.Width - buffer) Then
                            fcnt += 1

                        End If
                        strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                                 System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblproduct.Width - buffer) Then
                            fcnt += 1

                        End If

                    End If
                Next

            End If

            If fcnt = (doit * 3) Then
                If fs < 11.75 Then
                    fs = fs + 1
                End If





                For e As Integer = 0 To y.Controls.Count - 1
                    If y.Controls.Item(e).Name.Contains("pnl") Then
                        doit += 1

                        Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                        Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                        Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                        Dim strsize As New System.Drawing.SizeF()
                        strsize = Me.MeasureString(lbladdress.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                              System.Drawing.GraphicsUnit.Point, CType(0, Byte)))

                        If strsize.Width <= (lbladdress.Width - buffer) Then
                            fcnt += 1
                        Else
                            fcnt -= 1
                        End If
                        strsize = Me.MeasureString(lblname.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                           System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblname.Width - buffer) Then
                            fcnt += 1

                        End If
                        strsize = Me.MeasureString(lblproduct.Text, New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, _
                                                                                 System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
                        If strsize.Width <= (lblproduct.Width - buffer) Then
                            fcnt += 1

                        End If

                    End If
                Next
                If fcnt <> (doit * 3) Then
                    fcnt = 0
                    doit = 0
                    Exit Sub
                End If
            End If
            For e As Integer = 0 To y.Controls.Count - 1

                If y.Controls.Item(e).Name.Contains("pnl") Then


                    Dim lblname As Label = y.Controls.Item(e).Controls.Item(1)
                    Dim lbladdress As Label = y.Controls.Item(e).Controls.Item(3)
                    Dim lblproduct As Label = y.Controls.Item(e).Controls.Item(4)
                    Dim lnk As LinkLabel = y.Controls.Item(e).Controls.Item(2)
                    lblname.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lbladdress.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblproduct.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lnk.Font = New System.Drawing.Font("Verdana", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                End If
            Next
            fcnt = 0
            doit = 0
        End If
    End Sub
    Function day(ByVal dayofweek As Integer)
        Dim d As String
        Select Case dayofweek
            Case 0
                d = "Sunday"
            Case 1
                d = "Monday"
            Case 2
                d = "Tuesday"
            Case 3
                d = "Wednesday"
            Case 4
                d = "Thursday"
            Case 5
                d = "Friday"
            Case 6
                d = "Saturday"
        End Select

        Return d

    End Function

    Dim i As ListViewItem
    Private Sub Link(ByVal sender As Object, ByVal e As System.EventArgs)
        If Sales.cboSalesList.Text <> "Issue Leads List" Then
            Sales.cboSalesList.SelectedItem = "Issue Leads List"
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
    End Sub

    Private Sub Cbo(ByVal sender As Object, ByVal e As System.EventArgs)
       
        Dim combo As ComboBox = sender
        If combo.SelectedItem = "<Add New>" Then
            combo.SelectedItem = Nothing
            AddSalesRep.combo = combo
            AddSalesRep.ShowDialog()
            Dim r As Panel
            For Each r In Sales.pnlIssue.Controls
                Try
                    Dim panel As Panel = r
                    Dim cbox1 As ComboBox = r.Controls.Item(6)
                    Dim cbox2 As ComboBox = r.Controls.Item(7)
                    Dim Orep1 As String = cbox1.SelectedItem
                    Dim Orep2 As String = cbox2.SelectedItem
                    Pop_Reps(cbox1)
                    Pop_Reps(cbox2)
                    cbox1.SelectedItem = Orep1
                    cbox2.SelectedItem = Orep2
                    If cbox1.SelectedItem = Nothing And Orep1 <> "" Then
                        cbox1.Items.Add(Orep1)
                        cbox1.SelectedItem = Orep1
                    End If
                    If cbox2.SelectedItem = Nothing And Orep2 <> "" Then
                        cbox2.Items.Add(Orep2)
                        cbox2.SelectedItem = Orep2
                    End If
                Catch ex As Exception

                End Try

            Next

            Exit Sub
        ElseIf combo.SelectedItem = "_____________________________________" Then
            combo.SelectedItem = Nothing
            Exit Sub
        Else
            If combo.SelectedItem = "" And combo.Name.Contains("Rep1") Then
                Dim cbo2 As ComboBox = combo.Parent.Controls.Item(7)
                If cbo2.SelectedItem <> Nothing Or cbo2.SelectedItem <> "" Then
                    combo.SelectedItem = cbo2.SelectedItem
                    cbo2.SelectedItem = ""

                End If
            ElseIf combo.Name.Contains("Rep2") And combo.SelectedItem <> "" Then
                Dim cboo As ComboBox = combo.Parent.Controls.Item(6)
                If cboo.SelectedItem = "" Or cboo.SelectedItem = Nothing Then
                    cboo.SelectedItem = combo.SelectedItem
                    combo.SelectedItem = ""
                End If
            End If
            Dim column As String
            If combo.Name.Contains("Rep1") Then
                column = "rep1"

            Else
                column = "rep2"

            End If
            Try
                Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                cnn.Open()
                Dim cmdPop As SqlCommand = New SqlCommand("update enterlead set " & column & " = '" & combo.SelectedItem & "', lastupdated = 'true' where id = '" & combo.Tag & "'", cnn)
                Dim r As SqlDataReader = cmdPop.ExecuteReader
                r.Read()
                r.Close()
                cnn.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Labels(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        Dim lbl As Control = sender

        PanelControl(lbl.Parent, e)
    End Sub

    Private Sub PanelControl(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'MsgBox(sender.ToString)

        '' check here to see if this lead is a called and cancelled.
        '' no matter the status, toggle the 'button called and cancelled' to its
        '' respective text 
        ''
        '' per andy 8-26-2015
        ''
        '' "Logged As Called and Cancelled"
        '' "Undo Called and Cancelled
        ''

        '' first find the 'lnk#ID#' control
        '' then get its text
        '' then query against sql to find last marketing status [iss.dbo.enter lead]
        '' if res = called and cancelled
        '' form button = "undo c and c"
        '' elseif res <> called and cancelled
        '' form button = "Logged as c and c"
        '' end if 


        ''
        '' target control on form  ' sales '
        '' ' btnCCIssue ' 
        ''

        '' loop to find linklabel
        ''
        Dim y As Control
        Dim where As Panel = sender

        Dim proxyLeadNum As String = ""
        For Each y In where.Controls
            If TypeOf y Is LinkLabel Then
                proxyLeadNum = y.Text
            End If
        Next

        '' send lead num to sql to find marketing results
        '' 
        Dim Last_EL_Result As String = GetMResult(proxyLeadNum)

        If Last_EL_Result = "Called and Cancelled" Then
            'MsgBox("Undo Called and Cancelled" & vbCrLf & "LeadNum : " & proxyLeadNum & " | Result : " & Last_EL_Result, MsgBoxStyle.Information, "DEBUG INFO")
            Sales.btnCCIssue.Text = "Undo Called and Cancelled"
        ElseIf Last_EL_Result <> "Called and Cancelled" Then
            'MsgBox("Log Appt. as Called and Cancelled" & vbCrLf & "LeadNum : " & proxyLeadNum & " | Result : " & Last_EL_Result, MsgBoxStyle.Information, "DEBUG INFO")
            Sales.btnCCIssue.Text = "Log Appt. as Called and Cancelled"
        End If



        Dim m = sender.findform
        Dim who As Panel = sender

        Dim selected As Boolean
        If who.BorderStyle = BorderStyle.FixedSingle And e.Button = MouseButtons.Left Then
            selected = True
        Else
            selected = False
        End If





        Dim c As Integer = Sales.pnlIssue.Controls.Count
        Dim i As Integer
        For i = 1 To c
            Dim all As Panel = Sales.pnlIssue.Controls(i - 1)
            If all.BorderStyle = BorderStyle.FixedSingle Then
                all.BorderStyle = BorderStyle.None
            End If
        Next
        If selected = True Then
            who.BorderStyle = BorderStyle.None
            STATIC_VARIABLES.CurrentID = ""
        Else
            who.BorderStyle = BorderStyle.FixedSingle
            STATIC_VARIABLES.CurrentID = who.Controls.Item(2).Text
        End If

     

    End Sub

    Private Sub show_notes(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim pct As PictureBox = sender
        Dim tt As ToolTip
        If pct.Name.Contains("pctIcon") Then
            tt = ttStatus
        ElseIf pct.Name.Contains("pctMN") Then
            tt = ttNotes2
        ElseIf pct.Name.Contains("pctSI") Then
            tt = ttNotes
        Else
            Exit Sub
        End If
        If tt.Active = False Then
            tt.Active = True
            tt.Show(tt.GetToolTip(sender), sender)
        ElseIf tt.Active = True Then
            tt.Active = False
        End If


    End Sub
    Private Sub scroll(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.y.Select()
    End Sub
    Private Sub Add_Rep(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim p As Panel = sender.parent
        Dim i As PictureBox = p.Controls.Item(8)
        If i.Tag = "minus" Then
            p.Size = New Size(p.Width, 48)
            p.Controls.Item(6).Location = New System.Drawing.Point(p.Controls.Item(6).Location.X, 14)
            p.Controls.Item(7).Location = New System.Drawing.Point(p.Controls.Item(7).Location.X, 14)
            p.Controls.Item(7).Visible = False
            i.Location = New System.Drawing.Point(i.Location.X, 16)
            i.Image = Sales.ImgIssued.Images(5)
            i.Tag = "Add"
            ttlabels.SetToolTip(i, "Add a 2nd Rep")
            Dim x
            For x = 0 To p.Controls.Count - 1
                Dim Cntl As Control = p.Controls.Item(x)
                If Cntl.Name.Contains("lbl") Then
                    Cntl.Size = New Size(Cntl.Width, p.Height)
                    Cntl.Location = New System.Drawing.Point(Cntl.Location.X, 0)
                End If
            Next
            Dim c As ComboBox = p.Controls.Item(7)
            If c.Text <> "" Then
                c.SelectedItem = Nothing
            End If
        Else
            p.Size = New Size(p.Width, 62)
            p.Controls.Item(6).Location = New System.Drawing.Point(p.Controls.Item(6).Location.X, 7)
            p.Controls.Item(7).Location = New System.Drawing.Point(p.Controls.Item(7).Location.X, 34)
            p.Controls.Item(7).Visible = True
            i.Location = New System.Drawing.Point(i.Location.X, 9)
            i.Image = Sales.ImgIssued.Images(7)
            i.Tag = "minus"
            ttlabels.SetToolTip(i, "Close 2nd Rep")
            Dim x
            For x = 0 To p.Controls.Count - 1
                Dim Cntl As Control = p.Controls.Item(x)
                If Cntl.Name.Contains("lbl") Then
                    Cntl.Size = New Size(Cntl.Width, p.Height)
                    Cntl.Location = New System.Drawing.Point(Cntl.Location.X, 0)
                End If
            Next
       
        End If
    End Sub
    Private Sub Pop_Reps(ByVal cbo As ComboBox)
        cbo.Items.Clear()
        ''
        '' edit 8-16-2015
        '' switching over to 'tblTestEmployee' and querying 'employee' and 'department' consolidation of tables and creation ONE structure for an 'employee' to be reusable'
        ''
        cbo.Items.Add("<Add New>")
        cbo.Items.Add("_____________________________________")
        cbo.Items.Add("")


        Try
            cnn2.Open()
            Dim cmdPop As SqlCommand = New SqlCommand(" Select Distinct FName, LName  from SalesRepPull order by lname asc", cnn2)
            Dim r As SqlDataReader = cmdPop.ExecuteReader
            While r.Read()
                cbo.Items.Add(r.Item(0) & " " & r.Item(1))
            End While
            r.Close()
            cnn2.Close()
        Catch ex As Exception

        End Try

    End Sub

#Region "Prop For ReadOnly List of Reps to not have to hit Dbase so much."
    '' 8-16-2015
    '' ad hoc
    '' not well designed. 
    '' garbage work around for new table structure 
    '' Wrong List Fucker   AG :>)
    Public Property ListOfReps As List(Of USER_LOGICv2.Employee)
        Get
            Return arReps
        End Get
        Set(value As List(Of USER_LOGICv2.Employee))
            arReps = value
        End Set
    End Property
#End Region


    Private Function MeasureString(ByVal text As String, ByVal font As Font) As SizeF

        Dim size As SizeF

        size = TextRenderer.MeasureText(text, font)

        Return size

    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


#Region "Get Marekting Result For btnCCIssue Toggle"
    '' edit 8-27-2015
    ''
    Private Function GetMResult(ByVal LeadNum As String)
        Dim res As String = ""
        Dim m_resCNX As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        m_resCNX.Open()
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT  MarketingResults FROM enterlead where id='" & LeadNum & "';", m_resCNX)
        res = cmdGET.ExecuteScalar
        m_resCNX.Close()
        m_resCNX = Nothing
        Return res.ToString
    End Function

#End Region




End Class

