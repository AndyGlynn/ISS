Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
imports System .Drawing .Graphics 

Public Class CustomerHistory



    
    Dim panel As Panel
    Private Department As String
    Dim ttStatus As New ToolTip
    Dim ttNotes As New ToolTip
    Private notes As String
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.CnnCustomerHistory)
 

    Public Sub SetUp(ByVal frm As Form, ByVal ID As Integer, ByVal TScboCustomerHistory As ToolStripComboBox)
        Department = TScboCustomerHistory.Text


        If Department = "All" Then
            Department = "%"
        ElseIf Department = "Marketing Department" Then

            Department = "Marketing"
        ElseIf Department = "Installation Department" Then
            Department = "Installation"
        ElseIf Department = "Finance Department" Then
            Department = "Financing"
        ElseIf Department = "Recovery Department" Then
            Department = "Recovery"
        ElseIf Department = "Sales Department" Then
            Department = "Sales"
        ElseIf Department = "Administration" Then
            Department = "Administration"
        ElseIf Department = "Phone Correspondence" Then
            Department = ""

        End If
        If frm.Name = "Confirming" Then
            panel = Confirming.pnlCustomerHistory
        ElseIf frm.Name = "WCaller" Then
            panel = WCaller.pnlCustomerHistory
        ElseIf frm.Name = "ConfirmingSingleRecord" Then
            panel = ConfirmingSingleRecord.pnlCustomerHistory
        ElseIf frm.Name = "Sales" Then
            panel = Sales.pnlCustomerHistory
        ElseIf frm.Name = "Administration" Then
            panel = Administration.pnlCustomerHistory
        ElseIf frm.Name = "Finance" Then
            panel = Finance.pnlCustomerHistory
        ElseIf frm.Name = "Installation" Then
            panel = Installation.pnlCustomerHistory
        ElseIf frm.Name = "MarketingManager" Then
            panel = MarketingManager.pnlCustomerHistory
        ElseIf frm.Name = "PreviousCustomer" Then
            panel = PreviousCustomer.pnlCustomerHistory
        ElseIf frm.Name = "Recovery" Then
            panel = Recovery.pnlCustomerHistory
        End If


        panel.Controls.Clear()
        panel.SuspendLayout()


        Dim cmdGet As SqlCommand = New SqlCommand("dbo.PopulateCustomerHistory", cnn)
        Dim cmdCNT As SqlCommand = New SqlCommand("Select Count (ID) from LeadHistory where Leadnum = " & ID, cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        Dim param2 As SqlParameter = New SqlParameter("@Department", Department)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cnn.Open()


        Dim r2 As SqlDataReader = cmdCNT.ExecuteReader
        r2.Read()
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
            p.Name = "pnl" & R1.Item(0)
            panel.Controls.Add(p)



            Dim pctDepartment As New PictureBox
            pctDepartment.Size = New System.Drawing.Size(18, 18)
            pctDepartment.Location = New System.Drawing.Point(3, 5)
            pctDepartment.Name = "pctDepartment" & R1.Item(0)

            Try
                pctDepartment.Tag = R1.Item(25)
            Catch ex As Exception
                pctDepartment.Tag = ""
            End Try
            pctDepartment.Visible = True

        



            'MsgBox(lblDash.Width.ToString)

            Dim pctSubDepartment As New PictureBox
            pctSubDepartment.Size = New System.Drawing.Size(18, 18)
            pctSubDepartment.Location = New System.Drawing.Point(3, 5)
            pctSubDepartment.Name = "pctSubDepartment" & R1.Item(0)
            Try
                If R1.Item(26) <> "" Then
                    pctSubDepartment.Tag = R1.Item(26)
                End If

            Catch ex As Exception
                pctSubDepartment.Tag = ""
            End Try
            'pctSubDepartment.Visible = True


            Dim lblDate As New Label
            Dim sp = Split(R1.Item(2), " ", 2)
            Dim dt = sp(0)
            lblDate.Text = dt
            lblDate.Name = "lblDate" & R1.Item(0)
            lblDate.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            lblDate.Location = New System.Drawing.Point(26, 5)
            lblDate.Size = New System.Drawing.Size(75, 16)
            lblDate.AutoEllipsis = True



            Dim lblDescription As New Label
            lblDescription.AutoSize = False
            lblDescription.TextAlign = ContentAlignment.MiddleLeft
            lblDescription.Text = R1.Item(23)
            lblDescription.Name = "lblDescription" & R1.Item(0)
            lblDescription.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            lblDescription.Location = New System.Drawing.Point(106, 5)
            lblDescription.Anchor = AnchorStyles.Left Or AnchorStyles.Right
            lblDescription.AutoEllipsis = True




            Dim pctNotes As New PictureBox
            pctNotes.Anchor = AnchorStyles.Right
            pctNotes.Size = New System.Drawing.Size(28, 28)
            'dim w = p.Width 
            pctNotes.Location = New System.Drawing.Point(panel.Width - 33, 5)
            pctNotes.Name = "pctNotes" & R1.Item(0)

            ttNotes.ToolTipIcon = ToolTipIcon.Info
            ttNotes.ToolTipTitle = "Customer History Notes"
            ttNotes.SetToolTip(pctNotes, R1.Item(24))
            ttNotes.AutoPopDelay = 30000
            ttNotes.InitialDelay = 1
            ttNotes.ReshowDelay = 1
            pctNotes.Cursor = Cursors.Hand
            pctNotes.Visible = False
            pctNotes.Tag = R1.Item(24)
            pctNotes.BringToFront()


            AddHandler pctNotes.Click, AddressOf ResetToolTip
            If R1.Item(24) <> "" Then
                Dim c As New TruncateNotes
                c.Truncate(R1.Item(24), pctNotes)
                notes = c.NewSTRING
                ttNotes.SetToolTip(pctNotes, notes)
            End If
            Dim strsize As New System.Drawing.SizeF()
           

            strsize = Me.MeasureString(R1.Item(23).ToString, New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
            Dim z As Integer = 177
            Dim n As Integer
            Dim x As Integer
            Dim h As Integer = strsize.Width
            z = panel.Width - (lblDescription.Location.X + 43)
            

            If h >= z Then
                x = 46
                n = 32
            Else
                x = 28
                n = 16
            End If
            p.Size = New Size(panel.Width, x)
            lblDescription.Size = New System.Drawing.Size(z, n)
            ''Add Controls
            p.Controls.Add(pctDepartment)
            Try
                If R1.Item(26) <> "" Then
                    'p.Controls.Add(lblDash)
                    p.Controls.Add(pctSubDepartment)
                    pctSubDepartment.BringToFront()
                    'pctDepartment.Visible = False
                End If
            Catch ex As Exception


            End Try
            p.Controls.Add(lblDate)
            p.Controls.Add(lblDescription)
            p.Controls.Add(pctNotes)
           
       
            AddHandler p.MouseEnter, AddressOf scroll
            AddHandler pctSubDepartment.MouseEnter, AddressOf Rollover
            AddHandler pctDepartment.MouseLeave, AddressOf leave
            'AddHandler pctDepartment.MouseHover, AddressOf Rollover
            'AddHandler pctSubDepartment.MouseLeave, AddressOf leave


            Dim tooltip As String


            '    ''Icon Logic SubDepartment 
     

            If pctSubDepartment.Tag = "LoggedCorrespondence" And pctDepartment.Tag <> "" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(5)
                tooltip = " - Phone Correspondence"
            ElseIf pctSubDepartment.Tag = "LoggedCorrespondence" And pctDepartment.Tag = "" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(5)
                pctDepartment.Image = Main.ilCustomerHistory.Images(5)
                Me.ttStatus.SetToolTip(pctDepartment, "Phone Correspondence")

            ElseIf pctSubDepartment.Tag = "Confirming" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(7)
                tooltip = " - Confirmed Appointment"
            ElseIf pctSubDepartment.Tag = "Kill" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(11)
                tooltip = " - Kill"
            ElseIf pctSubDepartment.Tag = "DonotCall" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(12)
                tooltip = " - Do Not Call"
            ElseIf pctSubDepartment.Tag = "DonotMail" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(12)
                tooltip = " - Do Not Mail"
            ElseIf pctSubDepartment.Tag = "DonotBoth" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(12)
                tooltip = " - Do Not Call or Mail"
            ElseIf pctSubDepartment.Tag = "Scheduled Action" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(13)
                tooltip = " - Pending Scheduled Action"
            ElseIf pctSubDepartment.Tag = "Scheduled Action Done" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(13)
                tooltip = " - Scheduled Action - Completed"
            ElseIf pctSubDepartment.Tag = "Set Appointment" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(9)
                tooltip = " - Set Appointment"
            ElseIf pctSubDepartment.Tag = "Moved Appointment" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(9)
                tooltip = " - Moved Appointment"
            ElseIf pctSubDepartment.Tag = "Sales" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(1)
                tooltip = " - Sales Result"
            ElseIf pctSubDepartment.Tag = "System" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(14)
                tooltip = " - System Changed"
            ElseIf pctSubDepartment.Tag = "Cancelled" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(15)
                tooltip = " - Called & Cancelled"
            ElseIf pctSubDepartment.Tag = "Generated" Then
                pctSubDepartment.Image = Main.ilCustomerHistory.Images(16)
                tooltip = " - New Appointment"
            End If

            '    ''Icon Logic Department 
            If pctDepartment.Tag = "Marketing" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(0)
                ttStatus.SetToolTip(pctDepartment, "Marketing Department" & tooltip)
            ElseIf pctDepartment.Tag = "Sales" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(1)
                ttStatus.SetToolTip(pctDepartment, "Sales Department" & tooltip)
            ElseIf pctDepartment.Tag = "Installation" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(3)
                ttStatus.SetToolTip(pctDepartment, "Installation Department" & tooltip)
            ElseIf pctDepartment.Tag = "Financing" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(2)
                ttStatus.SetToolTip(pctDepartment, "Financing Department" & tooltip)
            ElseIf pctDepartment.Tag = "Administration" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(6)
                ttStatus.SetToolTip(pctDepartment, "Administration" & tooltip)
            ElseIf pctDepartment.Tag = "Recovery" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(4)
                ttStatus.SetToolTip(pctDepartment, "Recovery Department" & tooltip)
            ElseIf pctDepartment.Tag = "PreviousCustomer" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(8)
                ttStatus.SetToolTip(pctDepartment, "Previous Customer" & tooltip)
            ElseIf pctDepartment.Tag = "WarmCalling" Then
                pctDepartment.Image = Main.ilCustomerHistory.Images(17)
                ttStatus.SetToolTip(pctDepartment, "Warm Calling" & tooltip)
         End If





            'ElseIf pctDepartment.Tag = "Marketing Sales" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(1)
            '    ttStatus.SetToolTip(pctDepartment, "Sales Department")
            'End If



            'ElseIf pctSubDepartment.Tag = "Scheduled Action" Then
            '    pctSubDepartment.Image = Main.ilCustomerHistory.Images(13)
            '    ttStatus.SetToolTip(pctSubDepartment, "Scheduled Action - Completed")
            'ElseIf pctDepartment.Tag = "Install Scheduled Action" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(13)
            '    ttStatus.SetToolTip(pctDepartment, "Installation Department" & vbCr & "Scheduled Action - Completed")
            'ElseIf pctDepartment.Tag = "Admin Scheduled Action" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(13)
            '    ttStatus.SetToolTip(pctDepartment, "Administration" & vbCr & "Scheduled Action - Completed")
            'ElseIf pctDepartment.Tag = "Finance Scheduled Action" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(13)
            '    ttStatus.SetToolTip(pctDepartment, "Finance Department" & vbCr & "Scheduled Action - Completed")
            'ElseIf pctDepartment.Tag = "Marketing Scheduled Action" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(13)
            '    ttStatus.SetToolTip(pctDepartment, "Marketing Department" & vbCr & "Scheduled Action - Completed")
            'ElseIf pctDepartment.Tag = "Sales Scheduled Action" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(13)
            '    ttStatus.SetToolTip(pctDepartment, "Sales Department" & vbCr & "Scheduled Action - Completed")
            'ElseIf pctDepartment.Tag = "Set Appointment" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(9)
            '    ttStatus.SetToolTip(pctDepartment, "Set Appointment")
            'ElseIf pctDepartment.Tag = "Moved Appointment" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(9)
            '    ttStatus.SetToolTip(pctDepartment, "Moved Appointment")
            'ElseIf pctDepartment.Tag = "Marketing System" Then
            '    ttStatus.SetToolTip(pctDepartment, "Marketing - System Changed")
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(14)
            'ElseIf pctDepartment.Tag = "Recovery Moved Appointment" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(4)
            '    ttStatus.SetToolTip(pctDepartment, "Recovery - Moved Appointment")
            'ElseIf pctDepartment.Tag = "Previous Customer Moved Appointment" Then
            '    pctDepartment.Image = Main.ilCustomerHistory.Images(8)
            '    ttStatus.SetToolTip(pctDepartment, "Previous Customer - Moved Appointment")
            'End If


            ''add precvious customer structure 

            If R1.Item(24) <> "" Then
                pctNotes.Image = Main.ilCustomerHistory.Images(10)
                pctNotes.Visible = True
            End If

        End While
        R1.Close()
        cnn.Close()
        panel.ResumeLayout(False)
        panel.PerformLayout()
    End Sub
    Private Sub scroll(ByVal sender As Object, ByVal e As System.EventArgs)
    Me.panel.Select()
    End Sub
    Private Sub ResetToolTip(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.ttNotes.Active = False Then
            Me.ttNotes.Active = True
        ElseIf Me.ttNotes.Active = True Then
            Me.ttNotes.Active = False
        End If
    End Sub
    Private Sub Rollover(ByVal sender As Object, ByVal e As System.EventArgs)

        sender.SendToBack()
        'sender.visible = False
        'Me.pctSubDepartment.Visible = False
    End Sub
    Private Sub leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.SendToBack()
    End Sub
    Public Function MeasureString(ByVal text As String, ByVal font As Font) As SizeF

        Dim size As SizeF

        size = TextRenderer.MeasureText(text, Font)

        Return size

    End Function
  

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


End Class


