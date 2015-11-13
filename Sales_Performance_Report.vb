Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Drawing.Graphics

Public Class Sales_Performance_Report
    Dim panel As Panel = Sales.pnlPerformanceReport
    Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Dim tt As New ToolTip
    Public Sub New()
        panel.Controls.Clear()
        panel.SuspendLayout()
        'MsgBox(Sales.dtpSummary.Value.ToString & " | " & Sales.dtpSummary2.Value.ToString)
        Dim dt1 As String = Date_Strip(Sales.dtpSummary.Value)
        Dim dt2 As String = Date_Strip(Sales.dtpSummary2.Value)
        Dim cmdGet As SqlCommand
        cmdGet = New SqlCommand("dbo.SalesPerformanceReport", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@Date1", dt2)
        Dim param2 As SqlParameter = New SqlParameter("@Date2", dt1)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        Dim h As Integer = 0
        Dim fs As Single = 9.75
        Dim j As Integer = 0
        Dim buffer As Integer = 25
        Dim x As Double = Sales.tpSummary.Width - (61 + buffer)
        x = x / 8
        Dim pix As Double = 0
        Dim y As Double = (x / 2) - 61
        If y > 5 And y < 15 Then
            fs = 10.75
            h = 10
            j = 20
            pix = 0.5
        ElseIf y >= 15 Then
            fs = 11.75
            h = 20
            j = 20
            pix = 1
        End If
        'MsgBox(fs.ToString)
        Try
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            While r1.Read

                Dim p As New Panel
                Dim sw As Boolean
                Dim o As Integer = CInt(r1.Item(0))
                If o Mod 2 = 0 Then
                    sw = True
                Else
                    sw = False
                End If

                If sw = False Then
                    p.BackColor = Color.Transparent
                Else
                    p.BackColor = Color.WhiteSmoke
                End If




                p.Dock = DockStyle.Top
                p.BorderStyle = BorderStyle.None

                p.Name = "pnl" & r1.Item(0)
                p.Size = New Size(panel.Width, 32)
                panel.Controls.Add(p)
                If p.Name = "pnl1" Then
                    Dim p3 As New Panel
                    p3.Dock = DockStyle.Top
                    p3.BorderStyle = BorderStyle.None
                    p3.BackColor = Color.Transparent
                    p3.Size = New Size(panel.Width, 10)
                    panel.Controls.Add(p3)
                    Dim p2 As New Panel
                    p2.Dock = DockStyle.Top
                    p2.BorderStyle = BorderStyle.None
                    p2.BackColor = Color.Gray
                    p2.Size = New Size(panel.Width, 2)
                    panel.Controls.Add(p2)
                    AddHandler p3.MouseEnter, AddressOf panels
         
                End If


                Dim lnk As New LinkLabel
                Dim lbl As New Label
                If r1.Item(0) <> 1 Then
                    lnk.Text = r1.Item(1)
                    lnk.Name = "lnk" & r1.Item(0)
                    lnk.Location = New System.Drawing.Point(7, 8 - pix)
                    lnk.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lnk.AutoSize = True
                    'lnk.Dock = DockStyle.Fill
                    'lnk.Size = New Size(44, 32)
                    lnk.TextAlign = ContentAlignment.MiddleCenter
                    'lnk.Anchor = AnchorStyles.Bottom
                    AddHandler lnk.Click, AddressOf Link
                Else

                    lbl.Text = r1.Item(1)
                    lbl.Name = "lnk" & r1.Item(0)
                    lbl.Location = New System.Drawing.Point(7, 8 - pix)
                    lbl.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lbl.AutoSize = True
                    lbl.TextAlign = ContentAlignment.MiddleCenter
                    AddHandler lbl.MouseEnter, AddressOf panels
                    'lbl.Anchor = AnchorStyles.Bottom
                    'lbl.Dock = DockStyle.Fill
                    'lbl.Size = New Size(((Sales.lblIssued.Location.X + 61) - 22) - 7, 32)
                End If


                'Dim b = Sales.lblIssued.Location.X() + (Sales.lblIssued.Size.Width / 2)
                Dim lblIssued As New Label
                lblIssued.Text = r1.Item(2)
                lblIssued.Name = "lblIssued" & r1.Item(0)
                lblIssued.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblIssued.AutoSize = False
                lblIssued.Size = New Size(44 + h, 32)
                lblIssued.TextAlign = ContentAlignment.MiddleCenter
                lblIssued.AutoEllipsis = True
                'lblIssued.Size = New Size(32, 16)
                lblIssued.Location = New System.Drawing.Point((Sales.lblIssued.Location.X + (Sales.lblIssued.Width / 2)) - (22 + h), 0) '132
                AddHandler lblIssued.MouseEnter, AddressOf panels


                Dim lblDemos As New Label
                lblDemos.Text = r1.Item(3)
                lblDemos.Name = "lblDemos" & r1.Item(0)
                lblDemos.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblDemos.Location = New System.Drawing.Point((Sales.lblDNS.Location.X + (Sales.lblDNS.Width / 2)) - (42 + h), 0) '192
                lblDemos.AutoSize = False
                lblDemos.Size = New Size(32 + h, 32)
                lblDemos.TextAlign = ContentAlignment.MiddleLeft
                AddHandler lblDemos.MouseEnter, AddressOf panels

                Dim lblDemosPer As New Label
                lblDemosPer.Text = r1.Item(4) & "%"
                lblDemosPer.Name = "lblDemosPer" & r1.Item(0)
                lblDemosPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblDemosPer.Size = New Size(46 + h, 32)
                lblDemosPer.TextAlign = ContentAlignment.MiddleRight
                lblDemosPer.Location = New System.Drawing.Point(lblDemos.Location.X + (32 + h), 0) '224
                lblDemosPer.AutoSize = False
                AddHandler lblDemosPer.MouseEnter, AddressOf panels


                Dim lblResets As New Label
                lblResets.Text = r1.Item(5)
                lblResets.Name = "lblResets" & r1.Item(0)
                lblResets.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblResets.Size = New Size(32 + h, 32)
                lblResets.Location = New System.Drawing.Point((Sales.lblResets.Location.X + (Sales.lblResets.Width / 2)) - (42 + h), 0) '418
                lblResets.AutoSize = False
                lblResets.TextAlign = ContentAlignment.MiddleLeft
                AddHandler lblResets.MouseEnter, AddressOf panels


                Dim lblResetsPer As New Label
                lblResetsPer.Text = r1.Item(6) & "%"
                lblResetsPer.Name = "lblResetsPer" & r1.Item(0)
                lblResetsPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblResetsPer.Size = New Size(46 + h, 32)
                lblResetsPer.Location = New System.Drawing.Point(lblResets.Location.X + (32 + h), 0) '450
                lblResetsPer.AutoSize = False
                lblResetsPer.TextAlign = ContentAlignment.MiddleRight
                AddHandler lblResetsPer.MouseEnter, AddressOf panels


                Dim lblNotHits As New Label
                lblNotHits.Text = r1.Item(7)
                lblNotHits.Name = "lblNotHits" & r1.Item(0)
                lblNotHits.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblNotHits.Size = New Size(32 + h, 32)
                lblNotHits.Location = New System.Drawing.Point((Sales.lblNH.Location.X + (Sales.lblNH.Width / 2)) - (42 + h), 0) '305
                lblNotHits.AutoSize = False
                lblNotHits.TextAlign = ContentAlignment.MiddleLeft
                AddHandler lblNotHits.MouseEnter, AddressOf panels

                Dim lblNotHitsPer As New Label
                lblNotHitsPer.Text = r1.Item(8) & "%"
                lblNotHitsPer.Name = "lblNotHitsPer" & r1.Item(0)
                lblNotHitsPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblNotHitsPer.Size = New Size(46 + h, 32)
                lblNotHitsPer.Location = New System.Drawing.Point(lblNotHits.Location.X + (32 + h), 0) '337
                lblNotHitsPer.AutoSize = False
                lblNotHitsPer.TextAlign = ContentAlignment.MiddleRight
                AddHandler lblNotHitsPer.MouseEnter, AddressOf panels


                Dim lblSales As New Label



                lblSales.Text = r1.Item(11)
                lblSales.Name = "lblSales" & r1.Item(0)
                lblSales.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblSales.Size = New Size(32 + h, 32)
                lblSales.Location = New System.Drawing.Point((Sales.lblsales.Location.X + (Sales.lblsales.Width / 2)) - (42 + h), 0) '664
                lblSales.AutoSize = False
                lblSales.TextAlign = ContentAlignment.MiddleLeft
                AddHandler lblSales.MouseEnter, AddressOf panels
                Dim lblSalesPer As New Label


                lblSalesPer.Text = r1.Item(12) & "%"

                lblSalesPer.Name = "lblSalesPer" & r1.Item(0)
                lblSalesPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblSalesPer.Size = New Size(46 + h, 32)
                lblSalesPer.Location = New System.Drawing.Point(lblSales.Location.X + (32 + h), 0) '676
                lblSalesPer.AutoSize = False
                lblSalesPer.TextAlign = ContentAlignment.MiddleRight
                AddHandler lblSalesPer.MouseEnter, AddressOf panels
     

                Dim lblNoDemos As New Label
                lblNoDemos.Text = r1.Item(9)
                lblNoDemos.Name = "lblNoDemos" & r1.Item(0)
                lblNoDemos.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblNoDemos.Size = New Size(32 + h, 32)
                lblNoDemos.Location = New System.Drawing.Point((Sales.lblND.Location.X + (Sales.lblND.Width / 2)) - (42 + h), 0) '534
                lblNoDemos.AutoSize = False
                lblNoDemos.TextAlign = ContentAlignment.MiddleLeft
                AddHandler lblNoDemos.MouseEnter, AddressOf panels

                Dim lblNoDemosPer As New Label
                lblNoDemosPer.Text = r1.Item(10) & "%"
                lblNoDemosPer.Name = "lblNoDemosPer" & r1.Item(0)
                lblNoDemosPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblNoDemosPer.Size = New Size(46 + h, 32)
                lblNoDemosPer.Location = New System.Drawing.Point(lblNoDemos.Location.X + (32 + h), 0) '566
                lblNoDemosPer.AutoSize = False
                lblNoDemosPer.TextAlign = ContentAlignment.MiddleRight
                AddHandler lblNoDemosPer.MouseEnter, AddressOf panels


                Dim lblRCancels As New Label
                lblRCancels.Text = r1.Item(13)
                lblRCancels.Name = "lblRCancels" & r1.Item(0)
                lblRCancels.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblRCancels.Size = New Size(32 + h, 32)
                lblRCancels.Location = New System.Drawing.Point((Sales.lblRC.Location.X + (Sales.lblRC.Width / 2)) - (42 + h), 0) '757
                lblRCancels.AutoSize = False
                lblRCancels.TextAlign = ContentAlignment.MiddleLeft
                AddHandler lblRCancels.MouseEnter, AddressOf panels

                Dim lblRCancelsPer As New Label
                lblRCancelsPer.Text = r1.Item(14) & "%"
                lblRCancelsPer.Name = "lblRCancelsPer" & r1.Item(0)
                lblRCancelsPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblRCancelsPer.Size = New Size(46 + h, 32)
                lblRCancelsPer.Location = New System.Drawing.Point(lblRCancels.Location.X + (32 + h), 0) '789
                lblRCancelsPer.AutoSize = False
                lblRCancelsPer.TextAlign = ContentAlignment.MiddleRight
                AddHandler lblRCancelsPer.MouseEnter, AddressOf panels


                Dim lblSold As New Label
                Dim str As String = r1.Item(15)
                Dim index As Integer = str.Length
                str = str.Substring(0, index - 5)
                index = str.Length
                If index > 3 And index <= 6 Then
                    Dim str2 As String = str.Substring(index - 3, 3)
                    str = str.Substring(0, index - 3)
                    str = str & "," & str2
                ElseIf index > 6 Then
                    Dim str2 As String = str.Substring(index - 3, 3)
                    Dim str3 As String = str.Substring(index - 6, 3)
                    str = str.Substring(0, index - 6)

                    str = str & "," & str3 & "," & str2

                End If

                lblSold.Text = "$" & str
                lblSold.Name = "lblSold" & r1.Item(0)
                lblSold.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                lblSold.Size = New Size(80 + j, 32)
                If j > 0 And y > 0 Then
                    lblSold.Location = New System.Drawing.Point(Sales.pnlPerformanceReport.Width - (lblSold.Width + 10), 0) '869
                Else
                    lblSold.Location = New System.Drawing.Point(870, 0) '869 
                End If


                lblSold.AutoSize = False

                lblSold.TextAlign = ContentAlignment.MiddleLeft
                AddHandler lblSold.MouseEnter, AddressOf panels


                If r1.Item(0) <> 1 Then
                    tt.SetToolTip(lblIssued, "Total Appointments Issued to this Sales Rep")
                    tt.SetToolTip(lblDemos, "Total # of Demo/No Sale Results for this Sales Rep")
                    tt.SetToolTip(lblResets, "Total # of Reset Results for this Sales Rep")
                    tt.SetToolTip(lblSales, "Total # of Sale Results for this Sales Rep")
                    tt.SetToolTip(lblSold, "Total Dollars Sold for this Sales Rep")
                    tt.SetToolTip(lblRCancels, "Total # of Recission Cancels for this Sales Rep")
                    tt.SetToolTip(lblNoDemos, "Total # of No Demo Results for this Sales Rep")
                    tt.SetToolTip(lblNotHits, "Total # of Not Hit Results for this Sales Rep")
                    tt.SetToolTip(lblRCancelsPer, "Total % of Recission Cancels for this Sales Rep")
                    tt.SetToolTip(lblNoDemosPer, "Total % of No Demo Results for this Sales Rep")
                    tt.SetToolTip(lblNotHitsPer, "Total % of Not Hit Results for this Sales Rep")
                    tt.SetToolTip(lblDemosPer, "Total % of Demo/No Sale Results for this Sales Rep")
                    tt.SetToolTip(lblResetsPer, "Total % of Reset Results for this Sales Rep")
                    tt.SetToolTip(lblSalesPer, "Total % of Sale Results for this Sales Rep")
                Else
                    lblIssued.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblResets.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblNotHits.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblNoDemos.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblDemos.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblRCancels.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblSales.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblSold.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblResetsPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblNotHitsPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblNoDemosPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblDemosPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblRCancelsPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblSalesPer.Font = New System.Drawing.Font("Tahoma", fs!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    tt.SetToolTip(lblIssued, "Total Appointments Issued")
                    tt.SetToolTip(lblDemos, "Total # of Demo/No Sale Results")
                    tt.SetToolTip(lblResets, "Total # of Reset Results")
                    tt.SetToolTip(lblSales, "Total # of Sale Results")
                    tt.SetToolTip(lblSold, "Total Dollars Sold")
                    tt.SetToolTip(lblRCancels, "Total # of Recission Cancels")
                    tt.SetToolTip(lblNoDemos, "Total # of No Demo Results")
                    tt.SetToolTip(lblNotHits, "Total # of Not Hit Results")
                    tt.SetToolTip(lblDemosPer, "Total % of Demo/No Sale Results")
                    tt.SetToolTip(lblResetsPer, "Total % of Reset Results")
                    tt.SetToolTip(lblSalesPer, "Total % of Sale Results")
                    tt.SetToolTip(lblRCancelsPer, "Total % of Recission Cancels")
                    tt.SetToolTip(lblNoDemosPer, "Total % of No Demo Results")
                    tt.SetToolTip(lblNotHitsPer, "Total % of Not Hit Results")
                    Dim a As Double = 100 - ((CType(r1.Item(16), Integer) * 100) / CType(r1.Item(2), Integer))



                    Sales.lblAccuracy.Text = "This report is " & Math.Round(a, 2) & "% accurate"
                End If
                If r1.Item(0) <> 1 Then
                    p.Controls.Add(lnk)
                Else
                    p.Controls.Add(lbl)
                End If

                p.Controls.Add(lblIssued)
                p.Controls.Add(lblResets)
                p.Controls.Add(lblResetsPer)
                p.Controls.Add(lblNotHits)
                p.Controls.Add(lblNotHitsPer)
                p.Controls.Add(lblNoDemos)
                p.Controls.Add(lblNoDemosPer)
                p.Controls.Add(lblDemos)
                p.Controls.Add(lblDemosPer)
                p.Controls.Add(lblRCancels)
                p.Controls.Add(lblRCancelsPer)
                p.Controls.Add(lblSales)
                p.Controls.Add(lblSalesPer)
                p.Controls.Add(lblSold)

                ''Realign controls to center of column header 
                'lblIssued.Location = New System.Drawing.Point((lblIssued.Location.X() - 3) - (lblIssued.Size.Width / 2), 8)
                'lblResets.Location = New System.Drawing.Point(lblResets.Location.X() - 16, 8)
                If CType(r1.Item(16), Integer) <> 0 And r1.Item(1) <> "Total" Then
                    'MsgBox(r1.Item(16).ToString)
                    lblIssued.Text = (CType(lblIssued.Text, Integer) + CType(r1.Item(16), Integer)).ToString
                    lnk.BackColor = Color.Yellow
                    tt.SetToolTip(lnk, "There are outstanding sales results for this" & vbCr & "rep causing inaccuracy in this report!")
                End If



            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            MsgBox("Lost Network Connection! Sales Performance Report" & ex.ToString, MsgBoxStyle.Critical, "Server not Available")

            cnn.Close()
        End Try

        If panel.Controls.Count = 0 Then
            Dim lblNoData As New Label
            'lblNoData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            '        Or System.Windows.Forms.AnchorStyles.Left) _
            '        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            lblNoData.AutoSize = True
            lblNoData.Font = New System.Drawing.Font("Verdana", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

            lblNoData.Name = "lblNoData"
            lblNoData.Size = New System.Drawing.Size(616, 25)
            lblNoData.Location = New System.Drawing.Point((panel.Width / 2) - 308, (panel.Height / 2) - 12)
            lblNoData.TabIndex = 204
            lblNoData.Text = "There is No Data to Report for The Date Range Provided!"
            panel.Controls.Add(lblNoData)
        End If

        panel.ResumeLayout(False)
        panel.PerformLayout()








    End Sub
    Private Sub panels(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Sales.pnlPerformanceReport.Select()
    End Sub
    Private Sub Link(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rep As String = sender.text
        MsgBox("This will eventually run a rep report for " & rep)
    End Sub
    Private Function Date_Strip(ByVal Dt As Date)
        Dim dts As String = Dt.ToString
        Dim x = InStr(dts, " ")
        If x <> 0 Then
            dts = Microsoft.VisualBasic.Left(dts, x - 1)
        End If

        dts = dts & " 12:00:00 AM"
        Return dts
    End Function
End Class
